using System;
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

            var created = DocxHelper.CreateTextElement(textElement, textElement.Parent, TextValue);

            Assert.AreEqual(WordMl.TextRunName, created.Name);
            Assert.AreEqual(TextValue, created.Value);
        }

        [TestMethod]
        public void TestCreateTextElementWithinBodyDefaultWrap()
        {
            var textElement = new XElement("Text");
            const string TextValue = "Some text";
            var body = new XElement(WordMl.BodyName, textElement);

            var created = DocxHelper.CreateTextElement(textElement, textElement.Parent, TextValue);

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

            var created = DocxHelper.CreateTextElement(textElement, textElement.Parent, TextValue, hyperLink);

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

            var created = DocxHelper.CreateTextElement(textElement, textElement.Parent, TextValue, WordMl.SmartTagName);

            Assert.AreEqual(WordMl.SmartTagName, created.Name);
            var createdTextRun = created.Element(WordMl.TextRunName);
            Assert.IsNotNull(createdTextRun);
            Assert.AreEqual(TextValue, createdTextRun.Value);
        }

        [TestMethod]
        public void TestCreateAltChunkElement()
        {
            const int ExpectedAltChunkId = 10;

            var actualAltChunkElement = DocxHelper.CreateAltChunkElement(ExpectedAltChunkId);

            Assert.IsNotNull(actualAltChunkElement);
            Assert.AreEqual(WordMl.AltChunkName, actualAltChunkElement.Name);
            var actualAltChunkIdAttribute = actualAltChunkElement.Attribute(RelationshipMl.IdName);
            Assert.IsNotNull(actualAltChunkIdAttribute);
            Assert.AreEqual(string.Format("altChunkId{0}", ExpectedAltChunkId), actualAltChunkIdAttribute.Value);
        }

        [TestMethod]
        public void TestAddTextElementWithEmptyStyle()
        {
            var textElement = new XElement("Text");
            const string TextValue = "Some text";
            var body = new XElement(WordMl.BodyName, textElement);

            var created = DocxHelper.CreateTextElement(textElement, textElement.Parent, TextValue, WordMl.SmartTagName, string.Empty);
            Console.WriteLine(created);
        }

        [TestMethod]
        public void TestAddTextElementWithStyle()
        {
            var textElement = new XElement(WordMl.SdtName,
                                           new XElement(WordMl.SdtPrName, new XElement(WordMl.SdtName, "TagWidthStyle")),
                                           new XElement(WordMl.SdtContentName, new XElement(WordMl.TextRunName, "text")));
            const string TextValue = "Text with style";
            const string expectedStyleId = "SomeStyle";
            var body = new XElement(WordMl.BodyName, textElement);

            var created = DocxHelper.CreateTextElement(textElement, textElement.Parent, TextValue, expectedStyleId);
            Console.WriteLine(created);
            var textRunProperties = created.Element(WordMl.TextRunName).Element(WordMl.TextRunPropertiesName);
            Assert.IsNotNull(textRunProperties);
        }

        [TestMethod]
        public void TestAddEmptyParagraphInTableCell()
        {
            const string ExpectedRsidP = "00FF12FF";
            const int ExpectedAltChunkId = 1;
            var tableRowElement = new XElement(WordMl.TableRowName,
                                               new XAttribute(WordMl.RsidRPropertiesName, ExpectedRsidP));
            var altChunkElement = DocxHelper.CreateAltChunkElement(ExpectedAltChunkId);
            var tableCellElement = new XElement(WordMl.TableCellName, altChunkElement);
            tableRowElement.Add(tableCellElement);
            var document = new XElement(WordMl.BodyName, new XElement(WordMl.TableName, tableRowElement));

            DocxHelper.AddEmptyParagraphInTableCell(altChunkElement);

            Assert.IsNotNull(document);
            var actualElementAfterAltChunk = altChunkElement.NextElement();
            Assert.IsNotNull(actualElementAfterAltChunk);
            Assert.AreEqual(WordMl.ParagraphName, actualElementAfterAltChunk.Name);
            Assert.AreEqual(ExpectedRsidP, actualElementAfterAltChunk.Attribute(WordMl.RsidPName).Value);
            var actualRsidRAttribute = actualElementAfterAltChunk.Attribute(WordMl.RsidRName);
            Assert.IsNotNull(actualRsidRAttribute);
            Assert.AreEqual(actualRsidRAttribute.Value.Length, 8);
            Assert.IsTrue(actualRsidRAttribute.Value.Take(2).All(c => c.Equals('0')));
        }
    }
}
