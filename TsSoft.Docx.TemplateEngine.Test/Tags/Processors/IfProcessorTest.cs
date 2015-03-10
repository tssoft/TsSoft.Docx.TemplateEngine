using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TsSoft.Docx.TemplateEngine.Test.Tags.Processors
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml.Linq;

    using TsSoft.Docx.TemplateEngine.Tags;
    using TsSoft.Docx.TemplateEngine.Tags.Processors;

    [TestClass]
    public class IfProcessorTest : BaseProcessorTest
    {
        private XElement paragraph;
        private XElement startIf;
        private XElement endIf;
        private XElement body;
        private XElement data;

        [TestInitialize]
        public void TestInitialize()
        {
            this.startIf = new XElement(
                WordMl.SdtName,
                new XElement(
                    WordMl.SdtPrName,
                    new XElement(
                        WordMl.TagName,
                        new XAttribute(WordMl.ValAttributeName, "if")),
                    new XElement(
                        WordMl.IdName,
                        new XAttribute(WordMl.ValAttributeName, "1"))));
            this.paragraph = new XElement(
                WordMl.ParagraphName,
                new XElement(WordMl.ParagraphPropertiesName),
                new XElement(
                    WordMl.TextRunName, new XElement(WordMl.TextRunPropertiesName), new XElement(WordMl.TextName, "test text")));
            this.endIf = new XElement(
                WordMl.SdtName,
                new XElement(
                    WordMl.SdtPrName,
                    new XElement(
                        WordMl.TagName,
                        new XAttribute(WordMl.ValAttributeName, "endif")),
                    new XElement(
                        WordMl.IdName,
                        new XAttribute(WordMl.ValAttributeName, "2"))));
            this.body = new XElement("Body", this.startIf, this.paragraph, this.endIf);

            const string TruthfulElement = "truthfulelement";
            const string FalsyElement = "falsyelement";
            const string NonBooleanElement = "nonbooleanelement";
            this.data = new XElement("Data", new XElement(TruthfulElement, true), new XElement(FalsyElement, false), new XElement(NonBooleanElement, new object()));
        }

        [TestMethod]
        public void TestProcessTrue()
        {
            var processor = this.MakeProcessor("TruthfulElement");

            Console.WriteLine(this.body);

            processor.Process();

            Assert.IsFalse(this.body.Elements(WordMl.SdtName).Any(element => element.IsSdt()));
            this.ValidateTagsRemoved(this.body);
            var actualParagraph = this.body.Element(WordMl.ParagraphName);

            Assert.AreEqual(this.paragraph, actualParagraph);

            Console.WriteLine(this.body);
        }

        [TestMethod]
        public void TestProcessTrueWithLock()
        {
            var processor = this.MakeProcessor("TruthfulElement");
            processor.CreateDynamicContentTags = true;

            Console.WriteLine(this.body);

            processor.Process();

            this.ValidateTagsRemoved(this.body);
            Assert.IsTrue(
                this.body.Elements(WordMl.SdtName)
                    .Any(
                        element =>
                        element.Element(WordMl.SdtPrName)
                               .Element(WordMl.TagName)
                               .Attribute(WordMl.ValAttributeName)
                               .Value.ToLower()
                               .Equals("dynamiccontent")));
            Assert.AreEqual(1, this.body.Elements(WordMl.SdtName).Count());
            var actualParagraph =
                this.body.Element(WordMl.SdtName).Element(WordMl.SdtContentName).Element(WordMl.ParagraphName);

            Assert.IsNotNull(actualParagraph);
            Assert.AreEqual("test text", actualParagraph.Value);

            Console.WriteLine(this.body);
        } 
        
        [TestMethod]
        public void TestProcessFalse()
        {
            var processor = this.MakeProcessor("FalsyElement");

            Console.WriteLine(this.body);

            processor.Process();

            Assert.IsFalse(this.body.Elements(WordMl.SdtName).Any(element => element.IsSdt()));
            this.ValidateTagsRemoved(this.body);
            var actualParagraph = this.body.Element(WordMl.ParagraphName);

            Assert.IsNull(actualParagraph);

            Console.WriteLine(this.body);
        }

        [TestMethod]
        public void TestProcessFalseWithLock()
        {
            var processor = this.MakeProcessor("FalsyElement");
            processor.CreateDynamicContentTags = true;

            Console.WriteLine(this.body);

            processor.Process();

            this.ValidateTagsRemoved(this.body);
            Assert.IsTrue(
                this.body.Elements(WordMl.SdtName)
                    .Any(
                        element =>
                        element.Element(WordMl.SdtPrName)
                               .Element(WordMl.TagName)
                               .Attribute(WordMl.ValAttributeName)
                               .Value.ToLower()
                               .Equals("dynamiccontent")));
            Assert.AreEqual(1, this.body.Elements(WordMl.SdtName).Count());
            var actualParagraph =
                this.body.Element(WordMl.SdtName).Element(WordMl.SdtContentName).Element(WordMl.ParagraphName);

            Assert.IsNull(actualParagraph);

            Console.WriteLine(this.body);
        }

        [TestMethod]
        public void TestProcessNestedDynamicContent()
        {
            var dynamicContentTag = new XElement(WordMl.SdtName, new XElement(WordMl.SdtContentName, this.paragraph));
            this.startIf.Remove();
            this.endIf.Remove();
            this.body = new XElement("Body", this.startIf, dynamicContentTag, this.endIf);

            var processor = this.MakeProcessor("TruthfulElement");
            processor.CreateDynamicContentTags = true;
            processor.Tag.IfContent = new List<XElement> { dynamicContentTag };

            Console.WriteLine(this.body);

            processor.Process();

            this.ValidateTagsRemoved(this.body);
            Assert.IsTrue(
                this.body.Elements(WordMl.SdtName)
                    .Any(
                        element =>
                        element.Element(WordMl.SdtPrName)
                               .Element(WordMl.TagName)
                               .Attribute(WordMl.ValAttributeName)
                               .Value.ToLower()
                               .Equals("dynamiccontent")));
            Assert.AreEqual(1, this.body.Elements(WordMl.SdtName).Count());
            var actualParagraph =
                this.body.Element(WordMl.SdtName).Element(WordMl.SdtContentName).Element(WordMl.ParagraphName);

            Assert.IsNotNull(actualParagraph);
            Assert.AreEqual("test text", actualParagraph.Value);
        }
        
        [TestMethod]
        public void TestProcessNonBoolean()
        {
            var processor = this.MakeProcessor("NonBooleanElement");

            Console.WriteLine(this.body);

            processor.Process();

            this.ValidateTagsRemoved(this.body);
            var actualParagraph = this.body.Element(WordMl.ParagraphName);

            Assert.IsNull(actualParagraph);

            Console.WriteLine(this.body);
        }

        private IfProcessor MakeProcessor(string condition)
        {
            return new IfProcessor
                       {
                           DataReader = new DataReader(this.data),
                           Tag =
                               new IfTag
                                   {
                                       Conidition = condition,
                                       EndIf = this.endIf,
                                       IfContent = new List<XElement> { this.paragraph },
                                       StartIf = this.startIf,
                                   }
                       };
        }
    }
}
