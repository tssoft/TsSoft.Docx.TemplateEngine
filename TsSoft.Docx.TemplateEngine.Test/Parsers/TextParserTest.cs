using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Parsers;

namespace TsSoft.Docx.TemplateEngine.Test.Parsers
{
    [TestClass]
    public class TextParserTest //: BaseTagTest
    {
        [TestMethod]
        public void TestParser()
        {
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "TextParserTest.xml"))
            {
                var doc = XDocument.Load(docStream);
                var parser = new TextParser();
                var tag = parser.Do(doc.Descendants(WordMl.SdtName).First());
                Assert.IsNotNull(tag);
                Assert.AreEqual("//test/text", tag.Expression);
            }
        }
    }
}