using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags.Processors
{
    using System.Linq;

    [TestClass]
    public class TextProcessorTest : BaseProcessorTest
    {
        [TestMethod]
        public void TestProcess()
        {
            var textElement = new XElement("text");
            const string DynamicText = "Some people just want to watch the world burn.";
            textElement.Value = DynamicText;
            var data = new XElement("data", new XElement("test", textElement));

            var colorElement = new XElement(TextRunProperties.ColorName);
            colorElement.SetAttributeValue(WordMl.ValAttributeName, "FF0000");
            var textTag = new XElement(
                WordMl.SdtName,
                new XElement(
                    WordMl.SdtContentName,
                    new XElement(WordMl.TextRunName, new XElement(WordMl.TextRunPropertiesName, colorElement))),
                new XElement(WordMl.SdtPrName));
            var document = new XElement("body", textTag);

            var processor = new TextProcessor
            {
                DataReader = new DataReader(data),
                TextTag = new TextTag
                              {
                                  Expression = "//test/text",
                                  TagNode = textTag
                              }
            };

            processor.Process();

            this.ValidateTagsRemoved(document);
            Assert.IsNotNull(document.Descendants(WordMl.TextRunName).Single(d => d.Value == DynamicText));
            var colorProperty = document.Descendants(WordMl.TextRunName).Elements(WordMl.TextRunPropertiesName).Elements(TextRunProperties.ColorName).FirstOrDefault();
            Assert.IsNotNull(colorProperty);
            Assert.IsNotNull("FF0000", colorProperty.Attribute(WordMl.ValAttributeName).Value);
        }
    }
}