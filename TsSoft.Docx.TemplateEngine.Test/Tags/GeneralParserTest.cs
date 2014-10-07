using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    [TestClass]
    public class GeneralParserTest
    {


        private GeneralParser generalParser;
        private XDocument doc;

        private Mock<ITagProcessor> tagProcessorMock;
        private Mock<TextParser> textParserMock;
        private Mock<TableParser> tableParserMock;
        private Mock<RepeaterParser> repeaterParserMock;

        private Mock<RootProcessor> rootProcessorMock;
       
        [TestMethod]
        public void TestParse()
        {  
            generalParser = new GeneralParser();
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "table_document.xml");
            doc = XDocument.Load(docStream);

            tagProcessorMock = new Mock<ITagProcessor>();

            // Ожидаемое дерево процессов
            var expectedRootProcessor = new RootProcessor();
            expectedRootProcessor.AddProcessor(new TableProcessor());
            // --- 
            ITagProcessor actualProcessor;
           // rootProcessorMock.Setup(rp => rp.AddProcessor(It.IsAny<ITagProcessor>())).Callback(p => actualProcessor = p);
           //rootProcessorMock.Object.
        }

    }
}
