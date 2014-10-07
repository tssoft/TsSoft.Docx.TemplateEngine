using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TsSoft.Docx.TemplateEngine.Test
{
    using System.Linq;
    using System.Xml.Linq;

    using TsSoft.Docx.TemplateEngine.Tags.Processors;

    [TestClass]
    public class DocxHelperTest
    {
        [TestMethod]
        public void TestCreateTextElementWithinParagraph()
        {
            var textElement = new XElement("Text");
            const string TextValue = "Some text";
            var paragprahElement = new XElement(WordMl.ParagraphName, new XElement(WordMl.ParagraphPropertiesName), textElement);

            var created = DocxHelper.CreateTextElement(textElement.Parent, TextValue);

            Assert.AreEqual(WordMl.TextRunName, created.Name);
            Assert.AreEqual(TextValue, created.Value);
        }

        [TestMethod]
        public void TestCreateTextElementWithinBodyDefaultWrap()
        {
            var textElement = new XElement("Text");
            const string TextValue = "Some text";
            var body = new XElement(WordMl.BodyName, textElement);

            var created = DocxHelper.CreateTextElement(textElement.Parent, TextValue);

            Assert.AreEqual(WordMl.ParagraphName, created.Name);
            var createdTextRun = created.Element(WordMl.TextRunName);
            Assert.IsNotNull(createdTextRun);
            Assert.AreEqual(TextValue, createdTextRun.Value);
        }

        [TestMethod]
        public void TestCreateTextElementWithinBodyElementWrap()
        {
            var textElement = new XElement("Text");
            const string TextValue = "Some text";
            var body = new XElement(WordMl.BodyName, textElement);
            var hyperLink = new XElement(WordMl.HyperlinkName);

            var created = DocxHelper.CreateTextElement(textElement.Parent, TextValue, hyperLink);

            Assert.AreEqual(WordMl.HyperlinkName, created.Name);
            var createdTextRun = created.Element(WordMl.TextRunName);
            Assert.IsNotNull(createdTextRun);
            Assert.AreEqual(TextValue, createdTextRun.Value);
        }

        [TestMethod]
        public void TestCreateTextElementWithinBodyNamedWrap()
        {
            var textElement = new XElement("Text");
            const string TextValue = "Some text";
            var body = new XElement(WordMl.BodyName, textElement);

            var created = DocxHelper.CreateTextElement(textElement.Parent, TextValue, WordMl.SmartTagName);

            Assert.AreEqual(WordMl.SmartTagName, created.Name);
            var createdTextRun = created.Element(WordMl.TextRunName);
            Assert.IsNotNull(createdTextRun);
            Assert.AreEqual(TextValue, createdTextRun.Value);
        }
    }


}
