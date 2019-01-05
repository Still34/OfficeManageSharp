using CommandLine;

namespace OfficeManageSharp
{
    internal class Options
    {
        [Option("markFinal", HelpText = "Marks the file as final")]
        public bool ShouldMarkAsFinal { get; set; }

        [Option("removeFonts", HelpText = "Removes embedded fonts")]
        public bool ShouldRemoveEmbedFonts { get; set; }

        [Option('i', "input")] 
        public string InputDirectory { get; set; }

        [Option('r', "recurse", Default = false)]
        public bool IsRecursive { get; set; }

        [Option('s', "simulate", Default = false)]
        public bool IsSimulation { get; set; }
    }
}