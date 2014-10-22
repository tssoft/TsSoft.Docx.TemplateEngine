using System.Linq;
using CommandLine;
using CommandLine.Text;

namespace TsSoft.Docx.TemplateEngine.Demo
{
    /// <summary>
    /// Options for generating docx file
    /// </summary>
    internal class Options
    {
        [Option('t', "template", Required = true, HelpText = "File name of template *.docx word file")]
        public string TemplateFileName { get; set; }

        [Option('o', "output", Required = true, HelpText = "File name for saving generated *.docx document")]
        public string TargetFileName { get; set; }

        [Option('d', "data", Required = true, HelpText = "File name of *.xml data file")]
        public string DataFileName { get; set; }

        [Option('l', "log", HelpText = "File name of log file")]
        public string LogFileName { get; set; }

        [Option("append", HelpText = "Flag for append text to log file")]
        public bool LogAppend { get; set; }

        [Option("exception", MutuallyExclusiveSet = "ThrowException", HelpText = "Throw exception when data field missed in data file")]
        public bool ThrowException { get; set; }

        [Option("error", MutuallyExclusiveSet = "PrintError", HelpText = "Print error in output file for missed data field in data file")]
        public bool PrintError { get; set; }

        [Option("ignore", MutuallyExclusiveSet = "Ignore", HelpText = "Ignore data field when missed in data file")]
        public bool Ignore { get; set; }

        [Option("lock", HelpText = "Lock generated content in target document")]
        public bool LockDynamicContent { get; set; }

        [Option('h', "help", HelpText = "Display this help screen")]
        public bool Help { get; set; }

        [ParserState]
        public IParserState ParserState { get; set; }
        
        public string GetUsage()
        {
            var help = new HelpText();

            if (Help)
            {
                help.AddPreOptionsLine("Generate docx document from template and data entity");
                help.AddOptions(this);
            }
            else if (ParserState != null && ParserState.Errors.Any())
            {
                help.AddPreOptionsLine("ERROR(S):");
                help.AddPreOptionsLine(help.RenderParsingErrorsText(this, 2));
                help.AddPreOptionsLine("use key -h/--help for help");
            }

            return help;
        }
    }


}
