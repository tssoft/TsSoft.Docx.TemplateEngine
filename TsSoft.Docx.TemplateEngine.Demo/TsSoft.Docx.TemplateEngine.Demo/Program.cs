using CommandLine;
using System;
using System.IO;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new Options();
            if (!Parser.Default.ParseArguments(args, options) || options.Help)
            {
                Console.WriteLine(options.GetUsage());
            }
            else if (CanGenerate(options))
            {
                Generate(options);
            }
        }

        private static void Generate(Options options)
        {
            using (var templateStream = new FileStream(options.TemplateFileName, FileMode.Open))
            using (var destinationStream = new FileStream(options.TargetFileName, FileMode.Create))
            {
                var dataDocument = XDocument.Load(options.DataFileName);
                var generator = new DocxGenerator();
                try
                {
                    generator.GenerateDocx(templateStream, destinationStream, dataDocument);
                }
                catch (Exception exception)
                {
                    Console.WriteLine("Error: {0}", exception.Message);
                }
            }
        }

        private static bool CanGenerate(Options options)
        {
            if (!File.Exists(options.TemplateFileName))
            {
                Console.WriteLine("Template file {0} not found", options.TemplateFileName);
            }
            if (!File.Exists(options.DataFileName))
            {
                Console.WriteLine("Data file {0} not found", options.DataFileName);
            }

            return (File.Exists(options.TemplateFileName) && File.Exists(options.DataFileName));
        }
    }
}
