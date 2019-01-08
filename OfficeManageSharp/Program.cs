using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommandLine;
using Serilog;

namespace OfficeManageSharp
{
    internal class Program
    {
        private static async Task Main(string[] args)
        {
            InitLogger();
            await Parser.Default.ParseArguments<Options>(args)
                .MapResult(async x =>
                {
                    try
                    {
                        await ProcessDocumentsAsync(x);
                    }
                    catch (Exception e)
                    {
                        Log.Fatal(e, "An exception has occurred.");
                    }
                }, errors =>
                {
                    foreach (var error in errors) Log.Error(error.ToString());
                    return Task.CompletedTask;
                });
        }

        private static async Task ProcessDocumentsAsync(Options options)
        {
            Log.Information("Processing \"{path}\"...", options.InputDirectory);
            var docManager = new DocumentManager();
            var docs = GetDocuments(options.InputDirectory, options.IsRecursive, DocumentType.WordProcessor);
            if (docs?.Length > 0)
            {
                Log.Information("{docsCount} document(s) gathered from {input}...", docs.Length,
                    options.InputDirectory);
                if (options.ShouldMarkAsFinal)
                    await Task.WhenAll(docs.Select(s => Task.Run(() => docManager.MarkDocxAsFinal(s))));

                if (options.ShouldRemoveEmbedFonts)
                    await Task.WhenAll(docs.Select(s => Task.Run(() => docManager.RemoveDocxFonts(s))));
            }

            var slides = GetDocuments(options.InputDirectory, options.IsRecursive, DocumentType.Presentation);
            if (slides?.Length > 0)
            {
                Log.Information("{docsCount} presentation(s) gathered from {input}...", slides.Length,
                    options.InputDirectory);
                if (options.ShouldMarkAsFinal)
                    await Task.WhenAll(slides.Select(s => Task.Run(() => docManager.MarkPptxAsFinal(s))));
            }

            var spreadsheets = GetDocuments(options.InputDirectory, options.IsRecursive, DocumentType.Spreadsheet);
            if (spreadsheets?.Length > 0)
            {
                Log.Information("{docsCount} spreadsheet(s) gathered from {input}...", spreadsheets.Length,
                    options.InputDirectory);
                if (options.ShouldMarkAsFinal)
                    await Task.WhenAll(spreadsheets.Select(s => Task.Run(() => docManager.MarkXlsxAsFinal(s))));
            }
        }

        private static string[] GetDocuments(string filepath, bool isRecursive, DocumentType docType)
        {
            if (filepath == null) throw new ArgumentNullException(nameof(filepath));
            string searchPattern;
            switch (docType)
            {
                case DocumentType.WordProcessor:
                    searchPattern = ".docx";
                    break;
                case DocumentType.Spreadsheet:
                    searchPattern = ".xlsx";
                    break;
                case DocumentType.Presentation:
                    searchPattern = ".pptx";
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(docType), docType, null);
            }

            if (File.Exists(filepath))
                return Path.GetExtension(filepath).IndexOf(searchPattern, StringComparison.OrdinalIgnoreCase) != -1
                    ? new[] {filepath}
                    : null;

            return Directory.GetFiles(filepath, $"*{searchPattern}",
                isRecursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);
        }

        private static void InitLogger()
            => Log.Logger = new LoggerConfiguration().WriteTo.Console().CreateLogger();
    }
}