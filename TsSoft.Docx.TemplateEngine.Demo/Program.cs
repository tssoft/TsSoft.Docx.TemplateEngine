using CommandLine;
using System;
using System.IO;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Demo
{
    public class Program
    {
        public static void Main(string[] args)
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
                    DocxGeneratorSettings settings = null;
                    if (options.ThrowException)
                    {
                        settings = new DocxGeneratorSettings { MissingDataMode = MissingDataMode.ThrowException };
                    }
                    else if (options.PrintError)
                    {
                        settings = new DocxGeneratorSettings { MissingDataMode = MissingDataMode.PrintError };
                    }
                    else if (options.Ignore)
                    {
                        settings = new DocxGeneratorSettings { MissingDataMode = MissingDataMode.Ignore };
                    }
                    
                    if (settings == null)
                    {
                        settings = new DocxGeneratorSettings
                                       {
                                           DynamicContentMode =
                                               options.LockDynamicContent
                                                   ? DynamicContentMode.Lock
                                                   : DynamicContentMode.NoLock
                                       };
                    }
                    else
                    {
                        settings.DynamicContentMode = options.LockDynamicContent
                                                          ? DynamicContentMode.Lock
                                                          : DynamicContentMode.NoLock;
                    }
                    generator.GenerateDocx(templateStream, destinationStream, dataDocument, settings);
                }
                catch (Exception exception)
                {
                    ErrorLog(exception, options);
                }
            }
        }

        private static bool CanGenerate(Options options)
        {
            bool result = true;

            if (!File.Exists(options.TemplateFileName))
            {
                Console.WriteLine("Template file {0} not found", options.TemplateFileName);
                result = false;
            }
            if (result && !File.Exists(options.DataFileName))
            {
                Console.WriteLine("Data file {0} not found", options.DataFileName);
                result = false;
            }
            if (result 
                && ((options.ThrowException && options.PrintError && options.Ignore)
                || (options.ThrowException && options.PrintError && !options.Ignore)
                || (options.ThrowException && !options.PrintError && options.Ignore)
                || (!options.ThrowException && options.PrintError && options.Ignore)))
            {
                Console.WriteLine("--ignore, --error and --exception are mutually exclusive flags");
                result = false;
            }

            return result;
        }

        private static void ErrorLog(Exception exception, Options options)
        {
            Console.WriteLine("Error: {0}", exception.Message);
            if (string.IsNullOrEmpty(options.LogFileName))
            {
                Console.WriteLine("Stack trace:");
                Console.WriteLine(exception.StackTrace);
            }
            else
            {
                var mode = options.LogAppend ? FileMode.Append : FileMode.Create;
                using (var fileStream = new FileStream(options.LogFileName, mode))
                using (var writer = new StreamWriter(fileStream))
                {
                    writer.WriteLine("Error: {0}", exception.Message);
                    writer.WriteLine("Stack trace:");
                    writer.WriteLine(exception.StackTrace);
                    writer.WriteLine();
                }
            }
        }
    }
}
