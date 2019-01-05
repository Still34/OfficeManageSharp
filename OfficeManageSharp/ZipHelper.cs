using System;
using System.IO.Compression;

namespace OfficeManageSharp
{
    internal static class ZipHelper
    {
        public static void DoRebuildWithoutFonts(string zipFileIn, string zipFileOut)
        {
            using (var inputZipFile = ZipFile.OpenRead(zipFileIn))
            using (var outputZipFile = ZipFile.Open(zipFileOut, ZipArchiveMode.Update))
            {
                foreach (var zipArchiveEntry in inputZipFile.Entries)
                {
                    if (zipArchiveEntry.FullName.Contains("fonts", StringComparison.OrdinalIgnoreCase)) continue;
                    var entry = outputZipFile.CreateEntry(zipArchiveEntry.FullName, CompressionLevel.Optimal);
                    var oldEntry = zipArchiveEntry.Open();
                    var newEntry = entry.Open();
                    oldEntry.CopyTo(newEntry);
                    newEntry.Close();
                }
            }
        }
    }
}