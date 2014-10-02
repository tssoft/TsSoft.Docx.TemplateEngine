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
                var parser = new TextParser();
                var processor = parser.Parse(doc.Descendants(WordMl.SdtName).First());
                Assert.IsNotNull(processor);
                Assert.IsInstanceOfType(processor, typeof(TextParser));
                var tag = (processor as TextProcessor).TextTag;
                Assert.IsNotNull(tag);
                Assert.AreEqual("//test/text", tag.Expression);
            }
        }
    }
}