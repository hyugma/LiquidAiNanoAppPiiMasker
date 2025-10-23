using System.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace LiquidAiNanoAppPiiMasker;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void Window_DragEnter(object sender, System.Windows.DragEventArgs e)
    {
        if (e.Data != null && e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
        {
            e.Effects = System.Windows.DragDropEffects.Copy;
        }
        else
        {
            e.Effects = System.Windows.DragDropEffects.None;
        }
        e.Handled = true;
    }

    private void Window_DragOver(object sender, System.Windows.DragEventArgs e)
    {
        Window_DragEnter(sender, e);
    }

    private void Window_Drop(object sender, System.Windows.DragEventArgs e)
    {
        try
        {
            if (e.Data == null || !e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop)) return;
            var dropped = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop, false)!;

            var files = ExpandToSupportedFiles(dropped).ToArray();
            if (files.Length == 0)
            {
                System.Windows.MessageBox.Show("No supported files found. Supported: .docx, .pptx, .xlsx, .txt, .md, .rtf", "PII Masker", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var app = System.Windows.Application.Current as App;
            app?.StartMaskingFilesFromUI(files);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show("Failed to accept dropped files.\n" + ex.Message, "PII Masker", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static IEnumerable<string> ExpandToSupportedFiles(IEnumerable<string> paths)
    {
        foreach (var p in paths)
        {
            if (string.IsNullOrWhiteSpace(p)) continue;
            if (Directory.Exists(p))
            {
                foreach (var f in Directory.EnumerateFiles(p, "*.*", SearchOption.AllDirectories))
                {
                    if (IsSupportedExtension(f)) yield return f;
                }
            }
            else if (File.Exists(p))
            {
                if (IsSupportedExtension(p)) yield return p;
            }
        }
    }

    private static bool IsSupportedExtension(string path)
    {
        var ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
        return ext is ".docx" or ".pptx" or ".xlsx" or ".txt" or ".md" or ".rtf";
    }
}
