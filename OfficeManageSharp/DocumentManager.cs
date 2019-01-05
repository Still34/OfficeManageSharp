using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using Serilog;

namespace OfficeManageSharp
{
    // Thanks to the current version of OpenXML, none of the documents share an interface.
    internal class DocumentManager
    {
        public void MarkXlsxAsFinal(string input, bool doSave)
        {
            if (!File.Exists(input)) throw new FileNotFoundException();
            using (var ppt = SpreadsheetDocument.Open(input, true))
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

                var coreProp = ppt.CoreFilePropertiesPart ?? ppt.AddCoreFilePropertiesPart();
                coreProp.OpenXmlPackage.PackageProperties.ContentStatus = "Final";
                if (doSave) ppt.Save();
                Log.Information("Marked {inputName} as final.", filename);
            }
        }

        public void MarkPptxAsFinal(string input, bool doSave)
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

                var coreProp = ppt.CoreFilePropertiesPart ?? ppt.AddCoreFilePropertiesPart();
                coreProp.OpenXmlPackage.PackageProperties.ContentStatus = "Final";
                if (doSave) ppt.Save();

                Log.Information("Marked {inputName} as final.", filename);
            }
        }

        public void MarkDocxAsFinal(string input, bool doSave)
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

                var coreProp = doc.CoreFilePropertiesPart ?? doc.AddCoreFilePropertiesPart();
                coreProp.OpenXmlPackage.PackageProperties.ContentStatus = "Final";

                if (doSave) props.Save();
                Log.Information("Marked {inputName} as final.", filename);
            }
        }

        public void RemoveDocxFonts(string input, bool doSave)
        {
            using (var doc = WordprocessingDocument.Open(input, true))
            {

                doc.MainDocumentPart.FontTablePart.Fonts = new Fonts();

                if (doSave) doc.Save();
            }

            var filename = Path.GetFileNameWithoutExtension(input);
            var newFilename = Path.GetFullPath(input).Replace(filename, $"{filename}_");
            ZipHelper.DoRebuildWithoutFonts(input, newFilename);
        }
    }
}