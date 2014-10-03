using System.IO;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test
{
    internal class A
    {
        public int MyProperty { get; set; }
    }

    [TestClass]
    public class DocxGeneratorTest
    {
        private DocxGenerator docxGenerator;
        private Mock<Stream> templateStreamMock;
        private Mock<Stream> outputStreamMock;
        private Mock<DocxPackage> docxPackageMock;
        private Mock<ITagParser> parserMock;
        private Mock<AbstractProcessor> processorMock;
        private Mock<DataReader> stringDataReaderMock;
        private Mock<DataReader> entityDataReaderMock;
        private Mock<DataReader> xDocumentDataReaderMock;

        private XElement root;

        [TestInitialize]
        public void TestInitialize()
        {
            outputStreamMock = new Mock<Stream>();
            docxPackageMock = new Mock<DocxPackage>();
            var xDocument = new XDocument();
            root = xDocument.Root;

            docxPackageMock.SetupGet(p => p.DocumentPartXml).Returns(xDocument);
            docxPackageMock.Setup(p => p.Load()).Verifiable();
            docxPackageMock.Setup(p => p.Save()).Verifiable();

            var packageFactoryMock = new Mock<IDocxPackageFactory>();
            packageFactoryMock.Setup(f => f.Create(outputStreamMock.Object)).Returns(docxPackageMock.Object);

            processorMock = new Mock<AbstractProcessor>();
            processorMock.Setup(p => p.Process()).Verifiable();
            processorMock.SetupSet(p => p.DataReader).Verifiable();
            var processorFactoryMock = new Mock<IProcessorFactory>();
            processorFactoryMock.Setup(f => f.Create()).Returns(processorMock.Object);

            parserMock = new Mock<ITagParser>();
            var parserFactoryMock = new Mock<IParserFactory>();
            parserFactoryMock.Setup(f => f.Create()).Returns(parserMock.Object);

            stringDataReaderMock = new Mock<DataReader>();
            entityDataReaderMock = new Mock<DataReader>();
            xDocumentDataReaderMock = new Mock<DataReader>();
            var dataReaderFactoryMock = new Mock<IDataReaderFactory>();
            dataReaderFactoryMock.Setup(f => f.CreateReader(It.IsAny<string>())).Returns(stringDataReaderMock.Object);
            dataReaderFactoryMock.Setup(f => f.CreateReader(It.IsAny<A>())).Returns(entityDataReaderMock.Object);
            dataReaderFactoryMock.Setup(f => f.CreateReader(It.IsAny<XDocument>())).Returns(xDocumentDataReaderMock.Object);

            docxGenerator = new DocxGenerator
            {
                DataReaderFactory = dataReaderFactoryMock.Object,
                PackageFactory = packageFactoryMock.Object,
                ParserFactory = parserFactoryMock.Object,
                ProcessorFactory = processorFactoryMock.Object
            };

            templateStreamMock = new Mock<Stream>();

        }

        [TestMethod]
        public void TestGenerateDocxString()
        {
            docxGenerator.GenerateDocx(templateStreamMock.Object, outputStreamMock.Object, "whatever");
            Assert(stringDataReaderMock);
        }

        [TestMethod]
        public void TestGenerateDocxXDocument()
        {
            docxGenerator.GenerateDocx(templateStreamMock.Object, outputStreamMock.Object, new XDocument());
            Assert(xDocumentDataReaderMock);
        }

        [TestMethod]
        public void TestGenerateDocxEntity()
        {
            docxGenerator.GenerateDocx(templateStreamMock.Object, outputStreamMock.Object, new A());
            Assert(entityDataReaderMock);
        }

        private void Assert(Mock<DataReader> dataReaderMock)
        {
            docxPackageMock.Verify(p => p.Load(), Times.Once);

            parserMock.Verify(p => p.Parse(It.IsAny<AbstractProcessor>(), It.IsAny<XElement>()), Times.Once);
            parserMock.Verify(p => p.Parse(processorMock.Object, root), Times.Once);

            processorMock.VerifySet(p => p.DataReader = dataReaderMock.Object, Times.Once);
            processorMock.Verify(p => p.Process(), Times.Once);

            docxPackageMock.Verify(p => p.Save(), Times.Once);
        }
    }
}