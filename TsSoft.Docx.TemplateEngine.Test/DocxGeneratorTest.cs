using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System.IO;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test
{
    using System.Linq;
    using System.Xml.Serialization;

    using TsSoft.Commons.Utils;

    [TestClass]
    public class DocxGeneratorTest
    {
        private DocxGenerator docxGenerator;
        private Mock<DocxPackage> docxPackageMock;
        private Mock<ITagParser> parserMock;
        private Mock<AbstractProcessor> processorMock;
        private Mock<DataReader> stringDataReaderMock;
        private Mock<DataReader> entityDataReaderMock;
        private Mock<DataReader> documentDataReaderMock;

        private Stream templateStream;

        private Stream outputStream;

        private XElement root;

        [TestMethod]
        public void StubbedTestGenerateDocxString()
        {
            this.InitializeStubbedExecution();
            this.docxGenerator.GenerateDocx(this.templateStream, this.outputStream, "whatever");
            this.MakeAssertions(this.stringDataReaderMock);
        }

        [TestMethod]
        public void StubbedTestGenerateDocxXDocument()
        {
            this.InitializeStubbedExecution();
            this.docxGenerator.GenerateDocx(this.templateStream, this.outputStream, new XDocument());
            this.MakeAssertions(this.documentDataReaderMock);
        }

        [TestMethod]
        public void StubbedTestGenerateDocxEntity()
        {
            this.InitializeStubbedExecution();
            this.docxGenerator.GenerateDocx(this.templateStream, this.outputStream, new A());
            this.MakeAssertions(this.entityDataReaderMock);
        }

        [TestMethod]
        public void TestActualGeneration()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageTest.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            const string DynamicText = "My anaconda don't";
            generator.GenerateDocx(
                input,
                output,
                new SomeEntityWrapper
                    {
                        Test = new SomeEntity
                            {
                                Text = DynamicText
                            }
                    });
            var package = new DocxPackage(output);
            package.Load();

            XDocument documentPartXml = package.DocumentPartXml;
            Assert.IsFalse(documentPartXml.Descendants(WordMl.SdtName).Any());
            Assert.IsNotNull(documentPartXml.Descendants(WordMl.TextRunName).Single(e => e.Value == DynamicText));
        }
        /*
        [TestMethod]
        public void TestActualGenerationJointActiviesOriginal()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "joint_activies_original.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "reportNote.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data);
            var package = new DocxPackage(output);
            package.Load();

            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());            
        }

        [TestMethod]
        public void TestActualGenerationJointActiviesVersion2()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "joint_activies_version2.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "reportNote.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data);
            var package = new DocxPackage(output);
            package.Load();
            
            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
        }
        */

        [TestMethod]
        public void TestActualGenerationRepeater()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "Repeater.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "data.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data);
            var package = new DocxPackage(output);
            package.Load();

            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
        }

        [TestMethod]
        public void TestActualGenerationRepeaterInIf()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "RepeaterInIf.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "dataRepeaterInIf.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data);
            var package = new DocxPackage(output);
            package.Load();

            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
        }

        [TestMethod]
        public void TestActualGenerationDoubleIfAndText()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "IfText2.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "data.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data);
            var package = new DocxPackage(output);
            package.Load();

            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
        }

        [TestMethod]
        public void TestActualGenerationItemsAfterEndContent()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "RepeaterItemsAfterEndContent.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "dataItemsAfterEndContent.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data
                );
            var package = new DocxPackage(output);
            package.Load();
            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
        }

        [TestMethod]
        public void TestActualGenerationIfWithParagraphs()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "if.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "data.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data
                );
            var package = new DocxPackage(output);
            package.Load();
            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
        }
        /*
        [TestMethod]
        public void TestActualGenerationTableInRepeater()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "TableInRepeater.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "data.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data);

            var package = new DocxPackage(output);
            package.Load();
            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
            
        }
         */

        [TestMethod]
        public void TestActualGenerationStaticTextAfterTag()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "corruptedDocxx.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "DemoData2.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data);

            var package = new DocxPackage(output);
            package.Load();
            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
        }

        [TestMethod]
        public void TestActualGenerationDoubleItemIfWithItemText()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "corruptedDoc.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "DemoData2.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data);

            var package = new DocxPackage(output);
            package.Load();
            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
        }

        [TestMethod]
        public void TestActualGenerationIfWithTableWithParagraphs()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "ifTtable1.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "data.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data);

            var package = new DocxPackage(output);
            package.Load();
            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
        }

        [TestMethod]
        public void TestActualGenerationRepeaterWithTextWithParagraphs()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "repeatertext1.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "data.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data);

            var package = new DocxPackage(output);
            package.Load();
            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
        }

        [TestMethod]
        public void TestActualGenerationDoubleIf()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "DoubleIf.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "data.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data);

            var package = new DocxPackage(output);
            package.Load();
            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
        }
 
        [TestMethod]
        public void TestActualGenerationDoubleRepeater()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "DoubleRepeater.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "data.xml");
            var data = XDocument.Load(dataStream);
            generator.GenerateDocx(
                input,
                output,
                data);

            var package = new DocxPackage(output);
            package.Load();

            Assert.IsFalse(package.DocumentPartXml.Descendants(WordMl.SdtName).Any());
        }


        private void InitializeStubbedExecution()
        {
            this.docxPackageMock = new Mock<DocxPackage>();
            var document = new XDocument();
            this.root = document.Root;

            this.docxPackageMock.SetupGet(p => p.DocumentPartXml).Returns(document);
            this.docxPackageMock.Setup(p => p.Load()).Verifiable();
            this.docxPackageMock.Setup(p => p.Save()).Verifiable();

            var packageFactoryMock = new Mock<IDocxPackageFactory>();
            packageFactoryMock.Setup(f => f.Create(It.IsAny<Stream>())).Returns(this.docxPackageMock.Object);

            this.processorMock = new Mock<AbstractProcessor>();
            this.processorMock.Setup(p => p.Process()).Verifiable();
            this.processorMock.SetupSet(processor => processor.DataReader = It.IsAny<DataReader>()).Verifiable();
            var processorFactoryMock = new Mock<IProcessorFactory>();
            processorFactoryMock.Setup(f => f.Create()).Returns(this.processorMock.Object);

            this.parserMock = new Mock<ITagParser>();
            var parserFactoryMock = new Mock<IParserFactory>();
            parserFactoryMock.Setup(f => f.Create()).Returns(this.parserMock.Object);

            this.stringDataReaderMock = new Mock<DataReader>();
            this.entityDataReaderMock = new Mock<DataReader>();
            this.documentDataReaderMock = new Mock<DataReader>();
            var dataReaderFactoryMock = new Mock<IDataReaderFactory>();
            dataReaderFactoryMock.Setup(f => f.CreateReader(It.IsAny<string>())).Returns(this.stringDataReaderMock.Object);
            dataReaderFactoryMock.Setup(f => f.CreateReader(It.IsAny<A>())).Returns(this.entityDataReaderMock.Object);
            dataReaderFactoryMock.Setup(f => f.CreateReader(It.IsAny<XDocument>())).Returns(this.documentDataReaderMock.Object);

            this.docxGenerator = new DocxGenerator
            {
                DataReaderFactory = dataReaderFactoryMock.Object,
                PackageFactory = packageFactoryMock.Object,
                ParserFactory = parserFactoryMock.Object,
                ProcessorFactory = processorFactoryMock.Object
            };

            this.templateStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageTest.docx");
            this.outputStream = new MemoryStream();
        }

        private void MakeAssertions(IMock<DataReader> dataReaderMock)
        {
            this.docxPackageMock.Verify(p => p.Load(), Times.Once);

            this.parserMock.Verify(p => p.Parse(It.IsAny<AbstractProcessor>(), It.IsAny<XElement>()), Times.Once);
            this.parserMock.Verify(p => p.Parse(this.processorMock.Object, this.root), Times.Once);

            this.processorMock.VerifySet(p => p.DataReader = It.IsAny<DataReader>(), Times.Once);
            this.processorMock.VerifySet(p => p.DataReader = dataReaderMock.Object, Times.Once);
            this.processorMock.Verify(p => p.Process(), Times.Once);

            this.docxPackageMock.Verify(p => p.Save(), Times.Once);
        }

        public class SomeEntity
        {
            [XmlElement("text")]
            public string Text { get; set; }
        }

        public class SomeEntityWrapper
        {
            [XmlElement("test")]
            public SomeEntity Test { get; set; }
        }

        internal class A
        {
            public int MyProperty { get; set; }
        }
    }
}