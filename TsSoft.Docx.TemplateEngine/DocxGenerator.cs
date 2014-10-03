using System.IO;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine
{
    public class DocxGenerator
    {
        private IDocxPackageFactory packageFactory;
        private IProcessorFactory processorFactory;
        private IParserFactory parserFactory;
        private IDataReaderFactory dataReaderFactory;

        internal IDocxPackageFactory PackageFactory { private get { return packageFactory ?? new DocxPackageFactory(); } set { packageFactory = value; } }
        internal IProcessorFactory ProcessorFactory { private get { return processorFactory ?? new RootProcessorFactory(); } set { processorFactory = value; } }
        internal IParserFactory ParserFactory { private get { return parserFactory ?? new GeneralParserFactory(); } set { parserFactory = value; } }
        internal IDataReaderFactory DataReaderFactory { private get { return dataReaderFactory ?? new DataReaderFactory(); } set { dataReaderFactory = value; } }

        public void GenerateDocx<E>(Stream templateStream, Stream outputStream, E dataEntity)
        {
            var reader = DataReaderFactory.CreateReader(dataEntity);
            GenerateDocx(templateStream, outputStream, reader);
        }

        public void GenerateDocx(Stream templateStream, Stream outputStream, string dataXml)
        {
            var reader = DataReaderFactory.CreateReader(dataXml);
            GenerateDocx(templateStream, outputStream, reader);
        }

        public void GenerateDocx(Stream templateStream, Stream outputStream, XDocument dataXml)
        {
            var reader = DataReaderFactory.CreateReader(dataXml);
            GenerateDocx(templateStream, outputStream, reader);
        }

        private void GenerateDocx(Stream templateStream, Stream outputStream, DataReader reader)
        {
            templateStream.Seek(0, SeekOrigin.Begin);
            templateStream.CopyTo(outputStream);

            var package = PackageFactory.Create(outputStream);
            package.Load();

            var parser = ParserFactory.Create();
            var rootProcessor = ProcessorFactory.Create();
            parser.Parse(rootProcessor, package.DocumentPartXml.Root);

            rootProcessor.DataReader = reader;
            rootProcessor.Process();

            package.Save();
        }


    }

    internal interface IDocxPackageFactory
    {
        DocxPackage Create(Stream outputStream);
    }

    internal interface IParserFactory
    {
        ITagParser Create();
    }

    internal interface IProcessorFactory
    {
        AbstractProcessor Create();
    } 
    
    internal interface IDataReaderFactory
    {
        DataReader CreateReader<E>(E dataEntity);

        DataReader CreateReader(string dataXml);

        DataReader CreateReader(XDocument dataDocument);
    }

    internal class DocxPackageFactory : IDocxPackageFactory
    {
        public DocxPackage Create(Stream outputStream)
        {
            return new DocxPackage(outputStream);
        }
    }

    internal class GeneralParserFactory : IParserFactory
    {
        public ITagParser Create()
        {
            return new GeneralParser();
        }
    }

    internal class RootProcessorFactory : IProcessorFactory
    {
        public AbstractProcessor Create()
        {
            return new RootProcessor();
        }
    }

    internal class RootProcessor : AbstractProcessor { }
}