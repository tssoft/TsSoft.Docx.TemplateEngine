using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    [TestClass]
    public class TextParserTest
    {
        [TestMethod]
        public void TestParse()
        {
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "TextParserTest.xml"))
            {
                var doc = XDocument.Load(docStream);
                var processorMock = new TagProcessorMock<TextProcessor>();
                var parser = new TextParser();
                parser.Parse(processorMock, doc.Descendants(WordMl.SdtName).First());
                var processor = processorMock.InnerProcessor;
                var tag = processor.TextTag;
                Assert.IsNotNull(tag);
                Assert.AreEqual("//test/text", tag.Expression);
            }
        }
    }
}