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
            var (word, powerPoint, excel) = GetDocuments(options.InputDirectory, options.IsRecursive);
            var docManager = new DocumentManager();
            if (word.Length > 0)
            {
                Log.Information("{docsCount} document(s) gathered from {input}...", word.Length, options.InputDirectory);

                if (options.ShouldMarkAsFinal)
                    await Task.WhenAll(word.Select(s => Task.Run(() => docManager.MarkDocxAsFinal(s))));

                if (options.ShouldRemoveEmbedFonts)
                    await Task.WhenAll(word.Select(s => Task.Run(() => docManager.RemoveDocxFonts(s))));
            }

            if (powerPoint.Length > 0)
            {
                Log.Information("{docsCount} presentation(s) gathered from {input}...", powerPoint.Length,
                    options.InputDirectory);
                if (options.ShouldMarkAsFinal)
                    await Task.WhenAll(powerPoint.Select(s => Task.Run(() => docManager.MarkPptxAsFinal(s))));
            }

            if (excel.Length > 0)
            {
                Log.Information("{docsCount} spreadsheet(s) gathered from {input}...", excel.Length, options.InputDirectory);
                if (options.ShouldMarkAsFinal)
                    await Task.WhenAll(excel.Select(s => Task.Run(() => docManager.MarkXlsxAsFinal(s))));
            }
        }

        private static (string[] Word, string[] PowerPoint, string[] Excel) GetDocuments(string filepath,
            bool isRecursive)
        {
            if (filepath == null) throw new ArgumentNullException(nameof(filepath));
            if (!Directory.Exists(filepath)) throw new DirectoryNotFoundException();
            var docs = Directory.GetFiles(filepath, "*.docx",
                isRecursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);
            var ppts = Directory.GetFiles(filepath, "*.pptx",
                isRecursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);
            var xls = Directory.GetFiles(filepath, "*.xlsx",
                isRecursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);
            return (docs, ppts, xls);
        }

        private static void InitLogger()
            => Log.Logger = new LoggerConfiguration().WriteTo.Console().CreateLogger();
    }
}