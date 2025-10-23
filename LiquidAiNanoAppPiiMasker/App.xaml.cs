using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using Serilog;
using WF = System.Windows.Forms;
using LiquidAiNanoAppPiiMasker.Services;

namespace LiquidAiNanoAppPiiMasker;

/// <summary>
/// App entry: hosts the system tray icon and context menu.
/// </summary>
public partial class App : System.Windows.Application
{
    private WF.NotifyIcon? _trayIcon;
    private WF.ContextMenuStrip? _menu;
    private WF.ToolStripMenuItem? _autoOpenMenuItem;
    private string _appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LiquidAiNanoAppPiiMasker");
    private string _logFilePath = string.Empty;
    private string _modelDir = string.Empty;
    private string _modelFilePath = string.Empty;
    private FileProcessor? _processor;

    // Default model file name guess; users can replace later if repository changes.
    private const string ModelRepoUrl = "https://huggingface.co/LiquidAI/LFM2-350M-PII-Extract-JP-GGUF/resolve/main/";
    private const string DefaultModelFileName = "LFM2-350M-PII-Extract-JP-Q4_K_M.gguf"; // Fallback guess; adjusted at runtime if needed.

    // Setting persisted in a simple text file (true/false). Default true.
    private bool _autoOpenAfterRedaction = true;
    private string _settingsFilePath = string.Empty;

    private void Application_Startup(object sender, StartupEventArgs e)
    {
        try
        {
            Directory.CreateDirectory(_appDataDir);
            _logFilePath = Path.Combine(_appDataDir, "pii_masker_log.txt");
            _settingsFilePath = Path.Combine(_appDataDir, "settings.ini");
            _modelDir = Path.Combine(_appDataDir, "models");
            Directory.CreateDirectory(_modelDir);
            _modelFilePath = Path.Combine(_modelDir, DefaultModelFileName);

            ConfigureLogging();
            LoadSettings();

            Log.Information("Application starting up.");
            InitializeTrayIcon();

            // Hide any main window if created by template
            if (System.Windows.Application.Current != null && System.Windows.Application.Current.MainWindow != null)
            {
                System.Windows.Application.Current.MainWindow.Hide();
                System.Windows.Application.Current.MainWindow.ShowInTaskbar = false;
            }

            // Fire-and-forget model ensure workflow
            _ = Task.Run(EnsureModelAsync);
        }
        catch (Exception ex)
        {
            // As last resort, show a blocking message box and exit
            System.Windows.MessageBox.Show($"Startup error: {ex.Message}", "LiquidAiNanoAppPiiMasker", MessageBoxButton.OK, MessageBoxImage.Error);
            System.Windows.Application.Current.Shutdown(-1);
        }
    }

    private void Application_Exit(object sender, ExitEventArgs e)
    {
        try
        {
            if (_trayIcon != null)
            {
                _trayIcon.Visible = false;
                _trayIcon.Dispose();
            }
            Log.Information("Application exiting.");
        }
        catch { /* ignore */ }
        finally
        {
            Log.CloseAndFlush();
        }
    }

    private void ConfigureLogging()
    {
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.File(_logFilePath, rollingInterval: RollingInterval.Day, retainedFileCountLimit: 7, shared: true)
            .CreateLogger();
        Log.Information("Logging configured at {Path}", _logFilePath);
    }

    private void LoadSettings()
    {
        try
        {
            if (File.Exists(_settingsFilePath))
            {
                var text = File.ReadAllText(_settingsFilePath).Trim();
                if (bool.TryParse(text, out var parsed))
                {
                    _autoOpenAfterRedaction = parsed;
                }
            }
            else
            {
                SaveSettings(); // write defaults
            }
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Failed to load settings, using defaults.");
        }
    }

    private void SaveSettings()
    {
        try
        {
            File.WriteAllText(_settingsFilePath, _autoOpenAfterRedaction.ToString());
            Log.Information("Settings saved. AutoOpen={AutoOpen}", _autoOpenAfterRedaction);
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Failed to save settings");
        }
    }

    private void InitializeTrayIcon()
    {
        _menu = new WF.ContextMenuStrip();

        var redactFileItem = new WF.ToolStripMenuItem("Mask File...");
        redactFileItem.Click += (_, __) => OnRedactFile();

        var redactFolderItem = new WF.ToolStripMenuItem("Mask Folder...");
        redactFolderItem.Click += (_, __) => OnRedactFolder();

        _menu.Items.Add(redactFileItem);
        _menu.Items.Add(redactFolderItem);
        _menu.Items.Add(new WF.ToolStripSeparator());

        _autoOpenMenuItem = new WF.ToolStripMenuItem("Auto-open file after masking")
        {
            CheckOnClick = true,
            Checked = _autoOpenAfterRedaction
        };
        _autoOpenMenuItem.CheckedChanged += (_, __) =>
        {
            _autoOpenAfterRedaction = _autoOpenMenuItem.Checked;
            SaveSettings();
        };
        _menu.Items.Add(_autoOpenMenuItem);

        var showLogsItem = new WF.ToolStripMenuItem("Show Debug Logs...");
        showLogsItem.Click += (_, __) => OnShowLogs();
        _menu.Items.Add(showLogsItem);

        _menu.Items.Add(new WF.ToolStripSeparator());

        var exitItem = new WF.ToolStripMenuItem("Exit");
        exitItem.Click += (_, __) => System.Windows.Application.Current.Shutdown();
        _menu.Items.Add(exitItem);

        _trayIcon = new WF.NotifyIcon
        {
            Visible = true,
                Text = "LiquidAiNanoAppPiiMasker - PII Masker",
            ContextMenuStrip = _menu,
            Icon = TryLoadIcon()
        };
        _trayIcon.MouseClick += (_, args) =>
        {
            if (args.Button == WF.MouseButtons.Left)
            {
                ShowDropWindow();
            }
        };

        // Optional: Balloon tip on startup
        _trayIcon.BalloonTipTitle = "LiquidAiNanoAppPiiMasker";
        _trayIcon.BalloonTipText = "Running in the system tray. Right-click the icon for options.";
        _trayIcon.ShowBalloonTip(3000);

        Log.Information("Tray icon initialized.");
    }

    private void ShowDropWindow()
    {
        try
        {
            if (System.Windows.Application.Current == null) return;
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                var win = System.Windows.Application.Current.MainWindow as MainWindow;
                if (win == null || !win.IsLoaded)
                {
                    win = new MainWindow
                    {
                        ShowInTaskbar = true
                    };
                    System.Windows.Application.Current.MainWindow = win;
                    win.Show();
                }
                else
                {
                    if (win.WindowState == WindowState.Minimized) win.WindowState = WindowState.Normal;
                    win.Show();
                    win.Activate();
                    win.Topmost = true;
                    win.Topmost = false;
                    win.Focus();
                }
            });
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Failed to show drop window");
        }
    }

    public void StartMaskingFilesFromUI(IEnumerable<string> files)
    {
        try
        {
            EnsureProcessor();
            var arr = files as string[] ?? new List<string>(files).ToArray();
            _ = RunFileRedactionAsync(arr);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "StartMaskingFilesFromUI failed");
            ShowBalloon("Failed to start masking. See logs.");
        }
    }

    private System.Drawing.Icon TryLoadIcon()
    {
        try
        {
            var assetsDir = Path.Combine(AppContext.BaseDirectory, "Assets");

            // Prefer Liquid AI logo image if provided by user; overlay "PII".
            var logoCandidates = new[]
            {
                Path.Combine(assetsDir, "logo.png")
            };
            foreach (var p in logoCandidates)
            {
                if (File.Exists(p))
                {
                    using var logo = LoadLogoBitmap(p, 32, 32);
                    return CreatePiiIconFromBitmap(logo);
                }
            }

            // Fallback to a provided icon file if present
            var icoPath = Path.Combine(assetsDir, "icon.ico");
            if (File.Exists(icoPath))
            {
                using var fs = File.OpenRead(icoPath);
                return new System.Drawing.Icon(fs);
            }
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Failed to load custom logo/icon, generating PII icon.");
        }

        // Final fallback: generate a simple PII text icon
        return CreatePiiTextIcon(32, 32);
    }

    [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    private static System.Drawing.Icon CreatePiiIconFromBitmap(System.Drawing.Bitmap logo)
    {
        using var canvas = new System.Drawing.Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (var g = System.Drawing.Graphics.FromImage(canvas))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(System.Drawing.Color.Transparent);

            // Draw logo scaled to canvas
            g.DrawImage(logo, new System.Drawing.Rectangle(0, 0, 32, 32));

            // Bottom overlay bar
            using var overlay = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(190, 0, 0, 0));
            g.FillRectangle(overlay, new System.Drawing.Rectangle(0, 18, 32, 14));

            // "PII" label
            using var font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);
            var text = "PII";
            var size = g.MeasureString(text, font);
            float x = (32 - size.Width) / 2f;
            float y = 18 + (14 - size.Height) / 2f - 1f;
            using var brush = new System.Drawing.SolidBrush(System.Drawing.Color.White);
            g.DrawString(text, font, brush, x, y);
        }

        var hIcon = canvas.GetHicon();
        try
        {
            var icon = System.Drawing.Icon.FromHandle(hIcon);
            var clone = (System.Drawing.Icon)icon.Clone();
            DestroyIcon(hIcon);
            return clone;
        }
        catch
        {
            DestroyIcon(hIcon);
            throw;
        }
    }

    private static System.Drawing.Icon CreatePiiTextIcon(int width, int height)
    {
        using var bmp = new System.Drawing.Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (var g = System.Drawing.Graphics.FromImage(bmp))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(System.Drawing.Color.Transparent);

            // Simple circular badge background
            using var bg = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(0x2D, 0x2A, 0x8C)); // deep indigo
            g.FillEllipse(bg, 0, 0, width - 1, height - 1);
            using var pen = new System.Drawing.Pen(System.Drawing.Color.White, 1f);
            g.DrawEllipse(pen, 0.5f, 0.5f, width - 2, height - 2);

            // "PII"
            using var font = new System.Drawing.Font("Segoe UI", 11f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);
            var text = "PII";
            var size = g.MeasureString(text, font);
            float x = (width - size.Width) / 2f;
            float y = (height - size.Height) / 2f - 1f;
            using var brush = new System.Drawing.SolidBrush(System.Drawing.Color.White);
            g.DrawString(text, font, brush, x, y);
        }

        var hIcon = bmp.GetHicon();
        try
        {
            var icon = System.Drawing.Icon.FromHandle(hIcon);
            var clone = (System.Drawing.Icon)icon.Clone();
            DestroyIcon(hIcon);
            return clone;
        }
        catch
        {
            DestroyIcon(hIcon);
            throw;
        }
    }

    private static System.Drawing.Bitmap LoadLogoBitmap(string path, int width, int height)
    {
        using var src = new System.Drawing.Bitmap(path);
        var dest = new System.Drawing.Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var g = System.Drawing.Graphics.FromImage(dest);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
        g.Clear(System.Drawing.Color.Transparent);
        g.DrawImage(src, new System.Drawing.Rectangle(0, 0, width, height));
        return dest;
    }

    private async Task EnsureModelAsync()
    {
        try
        {
            if (!File.Exists(_modelFilePath))
            {
                Log.Information("Model not found. Starting first-run download...");
                ShowBalloon("Downloading PII model, please wait...");

                Directory.CreateDirectory(_modelDir);
                string url = ModelRepoUrl + DefaultModelFileName + "?download=true";

                using var http = new HttpClient();
                http.Timeout = TimeSpan.FromMinutes(30);

                using var response = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
                if (!response.IsSuccessStatusCode)
                {
                    Log.Warning("Default model file not available at {Url}. Status={Status}", url, response.StatusCode);
                    // If desired, try a simple fallback to the first .gguf known name list (not implemented in this MVP).
                    throw new InvalidOperationException("Failed to download model. You may need to place a .gguf model in the models folder manually.");
                }

                var tmpPath = _modelFilePath + ".downloading";
                using (var inStream = await response.Content.ReadAsStreamAsync())
                using (var outStream = File.Create(tmpPath))
                {
                    await CopyToAsyncWithProgress(inStream, outStream, response.Content.Headers.ContentLength, p =>
                    {
                        if (p % 10 == 0) // reduce log spam
                            Log.Information("Model download progress: {Progress}%", p);
                    }, CancellationToken.None);
                }
                File.Move(tmpPath, _modelFilePath, true);
                Log.Information("Model downloaded to {Path}", _modelFilePath);
                ShowBalloon("PII model downloaded successfully.");
            }
            else
            {
                Log.Information("Model already present at {Path}", _modelFilePath);
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Model download failed");
            ShowBalloon("Model download failed. Check logs from the tray menu.");
        }
    }

    private static async Task CopyToAsyncWithProgress(Stream source, Stream destination, long? totalBytes, Action<int> reportPercent, CancellationToken ct)
    {
        var buffer = new byte[81920];
        long totalRead = 0;
        int lastPercent = -1;
        while (true)
        {
            ct.ThrowIfCancellationRequested();
            int read = await source.ReadAsync(buffer.AsMemory(0, buffer.Length), ct);
            if (read == 0) break;
            await destination.WriteAsync(buffer.AsMemory(0, read), ct);
            totalRead += read;
            if (totalBytes.HasValue && totalBytes.Value > 0)
            {
                int percent = (int)(totalRead * 100 / totalBytes.Value);
                if (percent != lastPercent)
                {
                    lastPercent = percent;
                    reportPercent(percent);
                }
            }
        }
    }

    private void ShowBalloon(string message)
    {
        try
        {
            if (_trayIcon == null) return;
            _trayIcon.BalloonTipTitle = "LiquidAiNanoAppPiiMasker";
            _trayIcon.BalloonTipText = message;
            _trayIcon.ShowBalloonTip(3000);
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Failed to show balloon tip");
        }
    }


    private void OnShowLogs()
    {
        try
        {
            string dir = Path.GetDirectoryName(_logFilePath) ?? _appDataDir;
            string toOpen = _logFilePath;
            if (!File.Exists(toOpen))
            {
                var latest = Directory.EnumerateFiles(dir, "pii_masker_log*.txt")
                    .OrderByDescending(f => System.IO.File.GetLastWriteTime(f))
                    .FirstOrDefault();
                if (!string.IsNullOrEmpty(latest))
                {
                    toOpen = latest;
                }
            }

            Log.Information("Opening logs at {Path}", toOpen);
            Process.Start(new ProcessStartInfo
            {
                FileName = "notepad.exe",
                Arguments = $"\"{toOpen}\"",
                UseShellExecute = false
            });
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to open logs");
            System.Windows.MessageBox.Show($"Failed to open log file.\n{ex.Message}", "LiquidAiNanoAppPiiMasker", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnRedactFile()
    {
        try
        {
            using var ofd = new OpenFileDialogWpfWrapper
            {
                Filter = "Supported Files|*.docx;*.pptx;*.xlsx;*.txt;*.md;*.rtf",
                Multiselect = true,
                Title = "Select file(s) to mask"
            };
            if (ofd.ShowDialog() == true)
            {
                var files = ofd.FileNames;
                Log.Information("User selected {Count} files for masking.", files.Length);
                EnsureProcessor();
                _ = RunFileRedactionAsync(files);
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Mask File failed");
            ShowBalloon("Failed to start masking. See logs.");
        }
    }

    private void OnRedactFolder()
    {
        try
        {
            using var fbd = new WF.FolderBrowserDialog
            {
                Description = "Select a folder to mask (recursively)",
                ShowNewFolderButton = false
            };
            var result = fbd.ShowDialog();
            if (result == WF.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                var folder = fbd.SelectedPath;
                Log.Information("User selected folder for masking: {Folder}", folder);
                EnsureProcessor();
                _ = RunFolderRedactionAsync(folder);
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Mask Folder failed");
            ShowBalloon("Failed to start masking. See logs.");
        }
    }

    private void EnsureProcessor()
    {
        if (_processor == null)
        {
            _processor = new LiquidAiNanoAppPiiMasker.Services.FileProcessor(_modelFilePath, () => _autoOpenAfterRedaction);
        }
    }

    private async Task RunFileRedactionAsync(string[] files)
    {
        try
        {
            if (!File.Exists(_modelFilePath))
            {
                ShowBalloon("Model not found. Downloading...");
                await EnsureModelAsync();
                if (!File.Exists(_modelFilePath))
                {
                    ShowBalloon("Model unavailable. See logs.");
                    return;
                }
            }
            ShowBalloon($"Starting masking for {files.Length} file(s)...");
            await _processor!.ProcessFilesAsync(files);
            ShowBalloon($"Masking complete for {files.Length} file(s).");
        }
        catch (Exception ex)
        {
            Serilog.Log.Error(ex, "Masking run failed");
            ShowBalloon("Masking failed. See logs.");
        }
    }

    private async Task RunFolderRedactionAsync(string folder)
    {
        try
        {
            if (!File.Exists(_modelFilePath))
            {
                ShowBalloon("Model not found. Downloading...");
                await EnsureModelAsync();
                if (!File.Exists(_modelFilePath))
                {
                    ShowBalloon("Model unavailable. See logs.");
                    return;
                }
            }
            ShowBalloon("Starting masking in folder...");
            await _processor!.ProcessFolderAsync(folder);
            ShowBalloon("Folder masking complete.");
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Folder masking failed");
            ShowBalloon("Masking failed. See logs.");
        }
    }

    /// <summary>
    /// Simple WPF OpenFileDialog wrapper enabling using() pattern with Microsoft.Win32.OpenFileDialog
    /// </summary>
    private sealed class OpenFileDialogWpfWrapper : IDisposable
    {
        private readonly Microsoft.Win32.OpenFileDialog _dlg = new();

        public string Filter { get => _dlg.Filter; set => _dlg.Filter = value; }
        public bool Multiselect { get => _dlg.Multiselect; set => _dlg.Multiselect = value; }
        public string Title { get => _dlg.Title; set => _dlg.Title = value; }
        public string[] FileNames => _dlg.FileNames;
        public bool? ShowDialog() => _dlg.ShowDialog();
        public void Dispose() { /* nothing */ }
    }
}
