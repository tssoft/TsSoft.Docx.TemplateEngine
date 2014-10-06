using System.IO;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine
{
    using System.Diagnostics.CodeAnalysis;

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

    public class DocxGenerator
    {
        private IDocxPackageFactory packageFactory;
        private IProcessorFactory processorFactory;
        private IParserFactory parserFactory;
        private IDataReaderFactory dataReaderFactory;

        internal IDocxPackageFactory PackageFactory
        {
            private get
            {
                return this.packageFactory ?? new DocxPackageFactory();
            }
            set
            {
                this.packageFactory = value;
            }
        }

        internal IProcessorFactory ProcessorFactory
        {
            private get
            {
                return this.processorFactory ?? new RootProcessorFactory();
            }
            set
            {
                this.processorFactory = value;
            }
        }

        internal IParserFactory ParserFactory
        {
            private get
            {
                return this.parserFactory ?? new GeneralParserFactory();
            }
            set
            {
                this.parserFactory = value;
            }
        }

        internal IDataReaderFactory DataReaderFactory
        {
            private get
            {
                return this.dataReaderFactory ?? new DataReaderFactory();
            }
            set
            {
                this.dataReaderFactory = value;
            }
        }

        public void GenerateDocx<E>(Stream templateStream, Stream outputStream, E dataEntity)
        {
            var reader = this.DataReaderFactory.CreateReader(dataEntity);
            this.GenerateDocx(templateStream, outputStream, reader);
        }

        public void GenerateDocx(Stream templateStream, Stream outputStream, string dataXml)
        {
            var reader = this.DataReaderFactory.CreateReader(dataXml);
            this.GenerateDocx(templateStream, outputStream, reader);
        }

        public void GenerateDocx(Stream templateStream, Stream outputStream, XDocument dataXml)
        {
            var reader = this.DataReaderFactory.CreateReader(dataXml);
            this.GenerateDocx(templateStream, outputStream, reader);
        }

        private void GenerateDocx(Stream templateStream, Stream outputStream, DataReader reader)
        {
            templateStream.Seek(0, SeekOrigin.Begin);
            templateStream.CopyTo(outputStream);

            var package = this.PackageFactory.Create(outputStream);
            package.Load();

            var parser = this.ParserFactory.Create();
            var rootProcessor = this.ProcessorFactory.Create();
            parser.Parse(rootProcessor, package.DocumentPartXml.Root);

            rootProcessor.DataReader = reader;
            rootProcessor.Process();

            package.Save();
        }
    }

    [SuppressMessage("StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Reviewed. Suppression is OK here.")]
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

    internal class RootProcessor : AbstractProcessor
    {
    }
}