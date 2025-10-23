using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using LLama;
using LLama.Common;
using Serilog;

namespace LiquidAiNanoAppPiiMasker.Models
{
    public class PiiResult
    {
        public List<string> address { get; set; } = new();
        public List<string> company_name { get; set; } = new();
        public List<string> email_address { get; set; } = new();
        public List<string> human_name { get; set; } = new();
        public List<string> phone_number { get; set; } = new();

        public IEnumerable<string> AllValues()
        {
            foreach (var a in address) yield return a;
            foreach (var a in company_name) yield return a;
            foreach (var a in email_address) yield return a;
            foreach (var a in human_name) yield return a;
            foreach (var a in phone_number) yield return a;
        }
    }
}

namespace LiquidAiNanoAppPiiMasker.Services
{
    using LiquidAiNanoAppPiiMasker.Models;

    public class LlamaInference : IDisposable
    {
        private readonly string _modelPath;
        private readonly ModelParams _params;
        private readonly LLamaWeights _weights;
        private readonly StatelessExecutor _executor;

        public LlamaInference(string modelPath)
        {
            _modelPath = modelPath;
            _params = new ModelParams(_modelPath)
            {
                ContextSize = 4096,
                GpuLayerCount = 0
            };
            _weights = LLamaWeights.LoadFromFile(_params);
            _executor = new StatelessExecutor(_weights, _params);
        }

        public async Task<PiiResult> AnalyzeAsync(string text)
        {
            string prompt =
                "<|startoftext|><|im_start|>system\n" +
                "Extract <address>, <company_name>, <email_address>, <human_name>, <phone_number><|im_end|>\n" +
                "<|im_start|>user\n" +
                text +
                "<|im_end|>\n" +
                "<|im_start|>assistant\n";

            var inf = new InferenceParams
            {
                MaxTokens = 512,
                AntiPrompts = new List<string> { "<|im_end|>", "<|endoftext|>" }
            };

            var sb = new StringBuilder();
            await foreach (var token in _executor.InferAsync(prompt, inf))
            {
                sb.Append(token);
            }

            string raw = sb.ToString().Trim();
            Log.Information("Model raw output: {Output}", raw);

            string json = ExtractJson(raw);
            try
            {
                var result = JsonSerializer.Deserialize<PiiResult>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    AllowTrailingCommas = true
                });
                return result ?? new PiiResult();
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Failed to parse model JSON. Output: {Out}", raw);
                return new PiiResult();
            }
        }

        private static string ExtractJson(string s)
        {
            int start = s.IndexOf('{');
            int end = s.LastIndexOf('}');
            if (start >= 0 && end > start)
            {
                return s.Substring(start, end - start + 1);
            }
            return "{}";
        }

        public void Dispose()
        {
            _weights.Dispose();
        }
    }

#if false
    public record PdfBox(int PageNumber, double X, double Y, double Width, double Height);

    public static class PdfUtils
    {
        public static (string Text, List<(int Page, string Word, PdfBox Box)> Words) ExtractTextAndWordBoxes(string path)
        {
            using var doc = UglyToad.PdfPig.PdfDocument.Open(path);
            var allText = new StringBuilder();
            var results = new List<(int, string, PdfBox)>();
            int pageNum = 0;
            foreach (var page in doc.GetPages())
            {
                pageNum++;
                var letters = page.Letters;
                var words = NearestNeighbourWordExtractor.Instance.GetWords(letters);
                foreach (var w in words)
                {
                    var bbox = w.BoundingBox;
                    results.Add((pageNum, w.Text, new PdfBox(pageNum, bbox.Left, bbox.Bottom, bbox.Width, bbox.Height)));
                    allText.Append(w.Text);
                    allText.Append(' ');
                }
                allText.AppendLine();
            }
            return (allText.ToString(), results);
        }

        public static List<PdfBox> FindBoxesForString(string text, List<(int Page, string Word, PdfBox Box)> words, string phrase)
        {
            var boxes = new List<PdfBox>();
            if (string.IsNullOrWhiteSpace(phrase)) return boxes;

            var phraseTokens = phrase.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (phraseTokens.Length == 0) return boxes;

            int i = 0;
            while (i < words.Count)
            {
                if (string.Equals(words[i].Word, phraseTokens[0], StringComparison.OrdinalIgnoreCase))
                {
                    int j = 1;
                    int k = i + 1;
                    while (j < phraseTokens.Length && k < words.Count && words[k].Page == words[i].Page &&
                           string.Equals(words[k].Word, phraseTokens[j], StringComparison.OrdinalIgnoreCase))
                    {
                        j++; k++;
                    }
                    if (j == phraseTokens.Length)
                    {
                        var slice = words.Skip(i).Take(phraseTokens.Length).ToList();
                        double left = slice.First().Box.X;
                        double right = slice.Last().Box.X + slice.Last().Box.Width;
                        double bottom = slice.Min(w => w.Box.Y);
                        double top = slice.Max(w => w.Box.Y + w.Box.Height);
                        boxes.Add(new PdfBox(slice.First().Page, left, bottom, right - left, top - bottom));
                        i = k;
                        continue;
                    }
                }
                i++;
            }
            return boxes;
        }

        public static void RedactPdf(string inputPath, string outputPath, IEnumerable<PdfBox> boxes)
        {
            using var pdf = PdfReader.Open(inputPath, PdfDocumentOpenMode.Modify);
            var grouped = boxes.GroupBy(b => b.PageNumber);
            foreach (var g in grouped)
            {
                int pageIndex = g.Key - 1;
                if (pageIndex < 0 || pageIndex >= pdf.PageCount) continue;
                var page = pdf.Pages[pageIndex];
                using var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Prepend);
                double pageHeight = page.Height;
                var brush = XBrushes.Black;
                foreach (var b in g)
                {
                    double x = b.X;
                    double y = pageHeight - b.Y - b.Height; // convert from bottom-left to top-left origin
                    gfx.DrawRectangle(brush, new XRect(x, y, b.Width, b.Height));
                }
            }
            pdf.Save(outputPath);
        }
    }
#endif

    public static class OfficeUtils
    {
        public static string ExtractDocxText(string path)
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var sb = new StringBuilder();
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body != null)
            {
                foreach (var para in body.Descendants<Paragraph>())
                {
                    foreach (var text in para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                    {
                        sb.Append(text.Text);
                    }
                    sb.AppendLine();
                }
            }
            return sb.ToString();
        }

        public static string ExtractPptxText(string path)
        {
            using var doc = PresentationDocument.Open(path, false);
            var sb = new StringBuilder();
            var presPart = doc.PresentationPart;
            if (presPart != null)
            {
                foreach (var slidePart in presPart.SlideParts)
                {
                    foreach (var t in slidePart.Slide.Descendants<A.Text>())
                    {
                        if (!string.IsNullOrWhiteSpace(t.Text))
                        {
                            sb.Append(t.Text);
                            sb.Append(' ');
                        }
                    }
                    sb.AppendLine();
                }
            }
            return sb.ToString();
        }

        public static void RedactDocx(string inputPath, string outputPath, IEnumerable<string> piiValues)
        {
            File.Copy(inputPath, outputPath, overwrite: true);
            using var doc = WordprocessingDocument.Open(outputPath, true);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return;
            var list = piiValues.Distinct().OrderByDescending(s => s.Length).ToList();

            foreach (var t in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
            {
                if (string.IsNullOrEmpty(t.Text)) continue;
                string newText = t.Text;
                foreach (var p in list)
                {
                    if (string.IsNullOrEmpty(p)) continue;
                    string mask = MaskToBlocks(p);
                    newText = newText.Replace(p, mask, StringComparison.Ordinal);
                }
                if (newText != t.Text)
                {
                    t.Text = newText;
                }
            }
            doc.MainDocumentPart!.Document.Save();
        }

        public static void RedactPptx(string inputPath, string outputPath, IEnumerable<string> piiValues)
        {
            File.Copy(inputPath, outputPath, overwrite: true);
            using var doc = PresentationDocument.Open(outputPath, true);
            var presPart = doc.PresentationPart;
            if (presPart == null) return;
            var list = piiValues.Distinct().OrderByDescending(s => s.Length).ToList();

            foreach (var slidePart in presPart.SlideParts)
            {
                bool changed = false;
                foreach (var t in slidePart.Slide.Descendants<A.Text>())
                {
                    if (t.Text == null) continue;
                    string newText = t.Text;
                    foreach (var p in list)
                    {
                        if (string.IsNullOrEmpty(p)) continue;
                        string mask = MaskToBlocks(p);
                        newText = newText.Replace(p, mask, StringComparison.Ordinal);
                    }
                    if (newText != t.Text)
                    {
                        t.Text = newText;
                        changed = true;
                    }
                }
                if (changed)
                {
                    slidePart.Slide.Save();
                }
            }
        }

        public static string ExtractXlsxText(string path)
        {
            using var doc = SpreadsheetDocument.Open(path, false);
            var sb = new StringBuilder();
            var sst = doc.WorkbookPart!.SharedStringTablePart?.SharedStringTable;

            foreach (var wsPart in doc.WorkbookPart.WorksheetParts)
            {
                foreach (var cell in wsPart.Worksheet.Descendants<Cell>())
                {
                    string text = GetCellText(cell, sst);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        sb.Append(text);
                        sb.Append(' ');
                    }
                }
            }
            return sb.ToString();
        }

        public static void RedactXlsx(string inputPath, string outputPath, IEnumerable<string> piiValues)
        {
            File.Copy(inputPath, outputPath, overwrite: true);
            using var doc = SpreadsheetDocument.Open(outputPath, true);
            var sstPart = doc.WorkbookPart!.SharedStringTablePart;
            var sst = sstPart?.SharedStringTable;
            var list = piiValues.Distinct().ToList();

            foreach (var wsPart in doc.WorkbookPart.WorksheetParts)
            {
                bool changed = false;
                foreach (var cell in wsPart.Worksheet.Descendants<Cell>())
                {
                    string text = GetCellText(cell, sst);
                    if (string.IsNullOrEmpty(text)) continue;
                    foreach (var p in list)
                    {
                        if (text.Contains(p, StringComparison.Ordinal))
                        {
                            SetCellText(cell, "[MASKED]");
                            changed = true;
                            break;
                        }
                    }
                }
                if (changed)
                {
                    wsPart.Worksheet.Save();
                }
            }

            if (sstPart != null)
            {
                sstPart.SharedStringTable.Save();
            }
        }

        private static string GetCellText(Cell cell, DocumentFormat.OpenXml.Spreadsheet.SharedStringTable? sst)
        {
            if (cell == null) return "";
            string? v = cell.InnerText;

            if (cell.DataType != null && cell.DataType == CellValues.SharedString && int.TryParse(v, out int idx) && sst != null)
            {
                return sst.ElementAt(idx).InnerText;
            }
            return v ?? "";
        }

        private static void SetCellText(Cell cell, string text)
        {
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(text));
        }

        public static void RedactPlainTextFile(string inputPath, string outputPath, IEnumerable<string> piiValues)
        {
            string content = File.ReadAllText(inputPath);
            foreach (var p in piiValues.Distinct().OrderByDescending(s => s.Length))
            {
                if (string.IsNullOrEmpty(p)) continue;
                string mask = MaskToBlocks(p);
                content = content.Replace(p, mask, StringComparison.Ordinal);
            }
            File.WriteAllText(outputPath, content);
        }

        public static string MaskToBlocks(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            var chars = s.Select(c => char.IsWhiteSpace(c) ? c : '█').ToArray();
            return new string(chars);
        }
    }

    public class FileProcessor
    {
        private readonly string _modelPath;
        private readonly Func<bool> _autoOpenGetter;

        public FileProcessor(string modelPath, Func<bool> autoOpenGetter)
        {
            _modelPath = modelPath;
            _autoOpenGetter = autoOpenGetter;
        }

        public async Task ProcessFilesAsync(IEnumerable<string> files)
        {
            foreach (var f in files)
            {
                try
                {
                    await ProcessOneAsync(f);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "Failed processing file {File}", f);
                }
            }
        }

        public async Task ProcessFolderAsync(string folder)
        {
            var files = Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories)
                .Where(IsSupportedExtension)
                .ToList();
            Log.Information("Found {Count} files in folder {Folder}", files.Count, folder);
            await ProcessFilesAsync(files);
        }

        private static bool IsSupportedExtension(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            return ext is ".docx" or ".pptx" or ".xlsx" or ".txt" or ".md" or ".rtf";
        }

        private async Task ProcessOneAsync(string path)
        {
            Log.Information("Processing file: {File}", path);
            string ext = Path.GetExtension(path).ToLowerInvariant();
            string dir = Path.GetDirectoryName(path)!;
            string name = Path.GetFileNameWithoutExtension(path);
            string outPath = Path.Combine(dir, name + "_masked" + Path.GetExtension(path));

            string extracted = await ExtractTextAsync(path, ext);
            Log.Information("Text extracted. Length={Len}", extracted.Length);

            using var llama = new LlamaInference(_modelPath);
            var pii = await llama.AnalyzeAsync(extracted);
            Log.Information("PII detected. Counts: addr={A}, company={C}, email={E}, human={H}, phone={P}",
                pii.address.Count, pii.company_name.Count, pii.email_address.Count, pii.human_name.Count, pii.phone_number.Count);

            var all = pii.AllValues().Distinct().ToList();
            if (all.Count == 0)
            {
                Log.Information("No PII found. Copying file as masked without changes.");
                File.Copy(path, outPath, true);
                AfterOutput(outPath);
                return;
            }

            switch (ext)
            {
                case ".docx":
                    OfficeUtils.RedactDocx(path, outPath, all);
                    break;
                case ".pptx":
                    OfficeUtils.RedactPptx(path, outPath, all);
                    break;
                case ".xlsx":
                    OfficeUtils.RedactXlsx(path, outPath, all);
                    break;
                case ".txt":
                case ".md":
                case ".rtf":
                    OfficeUtils.RedactPlainTextFile(path, outPath, all);
                    break;
            }

            AfterOutput(outPath);
        }

        private async Task<string> ExtractTextAsync(string path, string ext)
        {
            switch (ext)
            {
                case ".docx":
                    return OfficeUtils.ExtractDocxText(path);
                case ".pptx":
                    return OfficeUtils.ExtractPptxText(path);
                case ".xlsx":
                    return OfficeUtils.ExtractXlsxText(path);
                case ".txt":
                case ".md":
                case ".rtf":
                    return await File.ReadAllTextAsync(path);
                default:
                    return "";
            }
        }


        private void AfterOutput(string outPath)
        {
            Log.Information("Masking complete: {Path}", outPath);
            try
            {
                if (_autoOpenGetter())
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = outPath,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Failed to auto-open output file");
            }
        }
    }
}
