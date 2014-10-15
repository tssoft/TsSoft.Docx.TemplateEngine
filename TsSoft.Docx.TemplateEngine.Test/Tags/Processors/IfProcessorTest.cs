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
    public class IfProcessorTest
    {
        private XElement paragraph;
        private XElement startIf;
        private XElement endIf;
        private XElement body;
        private XElement data;

        [TestInitialize]
        public void TestInitialize()
        {
            this.startIf = new XElement("If");
            this.paragraph = new XElement(
                WordMl.ParagraphName,
                new XElement(WordMl.ParagraphPropertiesName),
                new XElement(
                    WordMl.TextRunName, new XElement(WordMl.TextRunPropertiesName), new XElement(WordMl.TextName)));
            this.endIf = new XElement("EndIf");
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

            Assert.IsFalse(this.body.Descendants("If").Any());
            Assert.IsFalse(this.body.Descendants("EndIf").Any());
            var actualParagraph = this.body.Element(WordMl.ParagraphName);

            Assert.AreEqual(this.paragraph, actualParagraph);

            Console.WriteLine(this.body);
        }   
        
        [TestMethod]
        public void TestProcessFalse()
        {
            var processor = this.MakeProcessor("FalsyElement");

            Console.WriteLine(this.body);

            processor.Process();

            Assert.IsFalse(this.body.Descendants("If").Any());
            Assert.IsFalse(this.body.Descendants("EndIf").Any());
            var actualParagraph = this.body.Element(WordMl.ParagraphName);

            Assert.IsNull(actualParagraph);

            Console.WriteLine(this.body);
        }
        
        [TestMethod]
        public void TestProcessNonBoolean()
        {
            var processor = this.MakeProcessor("NonBooleanElement");

            Console.WriteLine(this.body);

            processor.Process();

            Assert.IsFalse(this.body.Descendants("If").Any());
            Assert.IsFalse(this.body.Descendants("EndIf").Any());
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
