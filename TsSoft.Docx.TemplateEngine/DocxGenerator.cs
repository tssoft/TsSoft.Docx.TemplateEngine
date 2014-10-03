using System.IO;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine
{
    public class DocxGenerator<E>
    {
        public void GenerateDocx(Stream templateStream, Stream outputStream, E dataEntity)
        {
            templateStream.Seek(0, SeekOrigin.Begin);
            templateStream.CopyTo(outputStream);

            var package = new DocxPackage(outputStream);
            package.Load();

            var parser = new GeneralParser();
            var rootProcessor = new RootProcessor();
            parser.Parse(rootProcessor, package.DocumentPartXml.Root);

            DataReader reader = DataReaderFactory.CreateReader<E>(dataEntity);
            rootProcessor.DataReader = reader;
            rootProcessor.Process();

            package.Save();
        }
    }

    internal class RootProcessor : AbstractProcessor { }
}