using System.IO;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TsSoft.Commons.Utils;

namespace TsSoft.Docx.TemplateEngine.Test
{
    [TestClass]
    public class TextTagInHeaderFooterAndDocumentTest
    {
        private readonly DocxPackage _package;

        public TextTagInHeaderFooterAndDocumentTest()
        {
            var input = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageHeaderFooterTest.docx");
            var output = new MemoryStream();
            var generator = new DocxGenerator();
            var stream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageHeaderFooterTest_data.xml");
            var data = XDocument.Load(stream);

            generator.GenerateDocx(input, output, data);

            _package = new DocxPackage(output).Load();
        }

        [TestMethod]
        public void TestTextTagsAreRemoved()
        {
            Assert.IsFalse(_package.Parts.SelectMany(x => x.PartXml.Descendants(WordMl.SdtName)).Any());
        }

        [TestMethod]
        public void TestTextTagsAreReplacedWithPopulatedRun()
        {
            Assert.AreEqual(3, _package.Parts.SelectMany(x => x.PartXml.Descendants(WordMl.TextRunName)).Count(e => e.Value == "test data"));
        }

        [TestMethod]
        public void TestParagraphsDoNotHaveNestedParagraphs()
        {
            var descendants = _package.Parts.SelectMany(x => x.PartXml.Descendants(WordMl.ParagraphName).Descendants());

            Assert.IsFalse(descendants.Any(el => el.Name == WordMl.ParagraphName));
        }
    }
}