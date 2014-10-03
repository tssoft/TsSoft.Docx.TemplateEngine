using System.IO;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine
{
    public class DocxGenerator<E>
    {
        public void GenerateDocx(Stream templateStream, Stream outputStream, E dataEntity)
        {
            var reader = DataReaderFactory.CreateReader<E>(dataEntity);
            GenerateDocx(templateStream, outputStream, reader);
        }

        public void GenerateDocx(Stream templateStream, Stream outputStream, string dataXml)
        {
            var reader = DataReaderFactory.CreateReader(dataXml);
            GenerateDocx(templateStream, outputStream, reader);
        }

        private void GenerateDocx(Stream templateStream, Stream outputStream, DataReader reader)
        {
            templateStream.Seek(0, SeekOrigin.Begin);
            templateStream.CopyTo(outputStream);

            var package = new DocxPackage(outputStream);
            package.Load();

            var parser = new GeneralParser();
            var rootProcessor = new RootProcessor();
            parser.Parse(rootProcessor, package.DocumentPartXml.Root);

            rootProcessor.DataReader = reader;
            rootProcessor.Process();

            package.Save();
        }
    }

    internal class RootProcessor : AbstractProcessor { }
}