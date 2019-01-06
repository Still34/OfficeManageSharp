using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using Serilog;

namespace OfficeManageSharp
{
    // Thanks to the current version of OpenXML, none of the documents share an interface.
    internal class DocumentManager
    {
        public void MarkXlsxAsFinal(string input)
        {
            if (!File.Exists(input)) throw new FileNotFoundException();
            using (var xls = SpreadsheetDocument.Open(input, true))
            {
                var filename = Path.GetFileName(input);
                Log.Information("Processing {inputFile}...", filename);
                var customProps = xls.CustomFilePropertiesPart;
                if (customProps == null)
                {
                    customProps = xls.AddCustomFilePropertiesPart();
                    customProps.Properties = new Properties();
                }

                var props = customProps.Properties;
                var markAsFinalProp = props
                    .OfType<CustomDocumentProperty>()
                    .FirstOrDefault(x => x.Name == "_MarkAsFinal");
                if (markAsFinalProp != null)
                {
                    Log.Warning("{inputFile} had already been marked as final, skipping...", filename);
                    return;
                }

                var newProp = new CustomDocumentProperty
                {
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                    Name = "_MarkAsFinal",
                    VTBool = new VTBool("true")
                };
                props.AppendChild(newProp);
                var pid = 2;
                foreach (var openXmlElement in props)
                    if (openXmlElement is CustomDocumentProperty customDocumentProperty)
                        customDocumentProperty.PropertyId = pid++;
                xls.PackageProperties.ContentStatus = "Final";
                Log.Information("Marked {inputName} as final.", filename);
            }
        }

        public void MarkPptxAsFinal(string input)
        {
            if (!File.Exists(input)) throw new FileNotFoundException();
            using (var ppt = PresentationDocument.Open(input, true))
            {
                var filename = Path.GetFileName(input);
                Log.Information("Processing {inputFile}...", filename);
                var customProps = ppt.CustomFilePropertiesPart;
                if (customProps == null)
                {
                    customProps = ppt.AddCustomFilePropertiesPart();
                    customProps.Properties = new Properties();
                }

                var props = customProps.Properties;
                var markAsFinalProp = props
                    .OfType<CustomDocumentProperty>()
                    .FirstOrDefault(x => x.Name == "_MarkAsFinal");
                if (markAsFinalProp != null)
                {
                    Log.Warning("{inputFile} had already been marked as final, skipping...", filename);
                    return;
                }

                var newProp = new CustomDocumentProperty
                {
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                    Name = "_MarkAsFinal",
                    VTBool = new VTBool("true")
                };
                props.AppendChild(newProp);
                var pid = 2;
                foreach (var openXmlElement in props)
                    if (openXmlElement is CustomDocumentProperty customDocumentProperty)
                        customDocumentProperty.PropertyId = pid++;
                ppt.PackageProperties.ContentStatus = "Final";

                Log.Information("Marked {inputName} as final.", filename);
            }
        }

        public void MarkDocxAsFinal(string input)
        {
            if (!File.Exists(input)) throw new FileNotFoundException();
            using (var doc = WordprocessingDocument.Open(input, true))
            {
                var filename = Path.GetFileName(input);
                Log.Information("Processing {inputFile}...", filename);
                var customProps = doc.CustomFilePropertiesPart;
                if (customProps == null)
                {
                    customProps = doc.AddCustomFilePropertiesPart();
                    customProps.Properties = new Properties();
                }

                var props = customProps.Properties;
                var markAsFinalProp = props
                    .OfType<CustomDocumentProperty>()
                    .FirstOrDefault(x => x.Name == "_MarkAsFinal");
                if (markAsFinalProp != null)
                {
                    Log.Warning("{inputFile} had already been marked as final, skipping...", filename);
                    return;
                }

                var newProp = new CustomDocumentProperty
                {
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                    Name = "_MarkAsFinal",
                    VTBool = new VTBool("true")
                };
                props.AppendChild(newProp);
                var pid = 2;
                foreach (var openXmlElement in props)
                    if (openXmlElement is CustomDocumentProperty customDocumentProperty)
                        customDocumentProperty.PropertyId = pid++;
                doc.PackageProperties.ContentStatus = "Final";

                Log.Information("Marked {inputName} as final.", filename);
            }
        }

        public void RemoveDocxFonts(string input)
        {
            using (var doc = WordprocessingDocument.Open(input, true))
            {
                var filename = Path.GetFileName(input);
                Log.Information("Removing embedded fonts from {inputFile}...", filename);
                var fontParts = doc.MainDocumentPart.Parts
                    .Where(x => x.OpenXmlPart.ContentType.Contains("font", StringComparison.OrdinalIgnoreCase))
                    .Select(x => x.OpenXmlPart)
                    .ToArray();
                if (fontParts.Any())
                    doc.MainDocumentPart.DeleteParts(fontParts);
                else
                    Log.Warning("No embedded fonts are found for {file}, skipping...", filename);
            }
        }
    }
}