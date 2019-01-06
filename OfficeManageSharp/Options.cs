using CommandLine;
using CommandLine.Text;

namespace OfficeManageSharp
{
    internal class Options
    {
        [Option('m', "markAsFinal", HelpText = "Marks the file as final")]
        public bool ShouldMarkAsFinal { get; set; }

        [Option('e', "removeFonts", HelpText = "Removes embedded fonts")]
        public bool ShouldRemoveEmbedFonts { get; set; }

        [Option('i', "input", Required = true)] 
        public string InputDirectory { get; set; }

        [Option('r', "recurse", Default = false)]
        public bool IsRecursive { get; set; }
    }
}