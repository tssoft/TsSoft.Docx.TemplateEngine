using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags.Processors
{
    [TestClass]
    public class RepeaterProcessorTest : BaseProcessorTest
    {
        private static Func<XElement, RepeaterElement> MakeElementCallback = element =>
            {
                var repeaterElement = new RepeaterElement
                                          {
                                              Elements = element.Elements().Select(MakeElementCallback),
                                              IsIndex = element.Name == "index",
                                              IsItem = element.Name == "item",
                                              XElement = element
                                          };
                if (repeaterElement.IsItem)
                {
                    repeaterElement.Expression = element.Value;
                }
                return repeaterElement;
            };

        private XElement startRepeater;

        private XElement endRepeater;

        private Mock<DataReader> dataReader;

        private IList<string> dateValues;

        private IList<string> subjectValues;

        private XElement body;

        private string xPath;

        private string staticText;

        private string firstLevelStaticText;

        private string secondLevelStaticText;

        private string thirdLevelStaticText;

        private string wrapper;

        [TestInitialize]
        public void Initialize()
        {
            this.startRepeater = new XElement(
                WordMl.SdtName,
                new XElement(
                    WordMl.SdtPrName,
                    new XElement(WordMl.TagName, new XAttribute(WordMl.ValAttributeName, "repeater")),
                    new XElement(WordMl.IdName, new XAttribute(WordMl.ValAttributeName, "1"))),
                "//test/certificates");
            this.endRepeater = new XElement(
                WordMl.SdtName,
                new XElement(
                    WordMl.SdtPrName,
                    new XElement(WordMl.TagName, new XAttribute(WordMl.ValAttributeName, "EndRepeater")),
                    new XElement(WordMl.IdName, new XAttribute(WordMl.ValAttributeName, "2"))));

            const string Index = "index";
            var indexElement = new XElement(Index);

            this.thirdLevelStaticText = "Very static. Wow. Much text.";
            this.staticText = "statictext";
            var thirdLevelStaticElement = new XElement(this.staticText, this.thirdLevelStaticText);

            this.firstLevelStaticText = "This text must render as it is";
            var firstLevelStaticElement = new XElement(this.staticText, this.firstLevelStaticText);

            this.secondLevelStaticText = "Just as its brother above, this text must render as it is, too";
            var secondLevelStaticElement = new XElement(this.staticText, this.secondLevelStaticText);

            var secondLevelElement = new XElement(WordMl.ParagraphName);
            var itemElement = new XElement("item");
            itemElement.Value = "./Date";
            secondLevelElement.Add(itemElement);
            secondLevelElement.Add(indexElement);
            secondLevelElement.Add(thirdLevelStaticElement);

            var firstLevelItem = new XElement("item");
            firstLevelItem.Value = "./Subject";

            this.wrapper = "wrapper";
            var firstLevelElement = new XElement("wrapper");
            firstLevelElement.Add(secondLevelElement);
            firstLevelElement.Add(secondLevelStaticElement);

            this.body = new XElement(
                "body", startRepeater, firstLevelItem, firstLevelElement, firstLevelStaticElement, endRepeater);

            Console.WriteLine(this.body.ToString());

            this.dateValues = new List<string> { "10.01.2014", "22.02.2014" };
            this.subjectValues = new List<string> { "Subject1", "Subject2" };

            var subject1 = new XElement("subject") { Value = this.subjectValues[0] };
            var date1 = new XElement("date") { Value = this.dateValues[0] };

            var subject2 = new XElement("subject") { Value = this.subjectValues[1] };
            var date2 = new XElement("date") { Value = this.dateValues[1] };

            var certificate1 = new XElement("certificate", subject1, date1);
            var certificate2 = new XElement("certificate", subject2, date2);

            this.dataReader = new Mock<DataReader>();

            this.xPath = "//test/certificates";
            this.dataReader.Setup(d => d.GetReaders(this.xPath))
                .Returns(() => new List<DataReader> { new DataReader(certificate1), new DataReader(certificate2) });
        }

        [TestMethod]
        public void TestProcess()
        {
            var processor = new RepeaterProcessor();
            processor.DataReader = this.dataReader.Object;
            processor.RepeaterTag = new RepeaterTag
                                        {
                                            Source = this.xPath,
                                            StartRepeater = this.startRepeater,
                                            EndRepeater = this.endRepeater,
                                            MakeElementCallback = MakeElementCallback
                                        };

            processor.Process();

            var dynamicContentTags =
                this.body.Elements(WordMl.SdtName)
                    .Where(
                        element =>
                        element.Element(WordMl.SdtPrName)
                               .Element(WordMl.TagName)
                               .Attribute(WordMl.ValAttributeName)
                               .Value.ToLower()
                               .Equals("dynamiccontent"));
            Assert.IsFalse(dynamicContentTags.Any());
            Console.WriteLine(body.ToString());

            this.CheckRepeaterContent(this.body);
        }

        [TestMethod]
        public void TestProcessWithLock()
        {
            var processor = new RepeaterProcessor();
            processor.DataReader = this.dataReader.Object;
            processor.RepeaterTag = new RepeaterTag
            {
                Source = this.xPath,
                StartRepeater = this.startRepeater,
                EndRepeater = this.endRepeater,
                MakeElementCallback = MakeElementCallback
            };
            processor.LockDynamicContent = true;

            processor.Process();

            var dynamicContentTags =
                this.body.Elements(WordMl.SdtName)
                    .Where(
                        element =>
                        element.Element(WordMl.SdtPrName)
                               .Element(WordMl.TagName)
                               .Attribute(WordMl.ValAttributeName)
                               .Value.ToLower()
                               .Equals("dynamiccontent"));
            Assert.IsTrue(dynamicContentTags.Any());
            Assert.AreEqual(1, dynamicContentTags.Count());

            this.CheckRepeaterContent(dynamicContentTags.First().Element(WordMl.SdtContentName));
        }

        private void CheckRepeaterContent(XElement container)
        {
            this.ValidateTagsRemoved(container);

            var subjects = container.Elements(WordMl.ParagraphName).ToList();
            Assert.AreEqual(6, container.Elements().Count());

            Assert.AreEqual(2, subjects.Count);
            Assert.AreEqual(this.subjectValues[0], subjects[0].Value);
            Assert.AreEqual(this.subjectValues[1], subjects[1].Value);

            var wrappers = container.Elements(this.wrapper).ToList();
            Assert.AreEqual(2, wrappers.Count);

            var secondLevelElements1 = wrappers[0].Elements().ToList();
            var secondLevelElements2 = wrappers[1].Elements().ToList();

            Assert.AreEqual(2, secondLevelElements1.Count);
            Assert.AreEqual(2, secondLevelElements2.Count);

            var secondLevelStaticText1 = secondLevelElements1[1];
            var secondLevelStaticText2 = secondLevelElements1[1];
            Assert.IsNotNull(secondLevelStaticText1);
            Assert.IsNotNull(secondLevelStaticText2);
            Assert.AreEqual(this.secondLevelStaticText, secondLevelStaticText1.Value);
            Assert.AreEqual(this.secondLevelStaticText, secondLevelStaticText2.Value);
            Assert.AreEqual(this.staticText, secondLevelStaticText1.Name);
            Assert.AreEqual(this.staticText, secondLevelStaticText2.Name);

            var paragraph1 = secondLevelElements1[0];
            var paragraph2 = secondLevelElements2[0];
            Assert.AreEqual(WordMl.ParagraphName, paragraph1.Name);
            Assert.AreEqual(WordMl.ParagraphName, paragraph2.Name);

            var paragraph1Elements = paragraph1.Elements().ToList();
            var paragraph2Elements = paragraph2.Elements().ToList();

            Assert.AreEqual(3, paragraph1Elements.Count);
            Assert.AreEqual(3, paragraph2Elements.Count);

            var paragraph1TextRuns = paragraph1Elements.Where(e => e.Name.Equals(WordMl.TextRunName)).ToList();
            var paragraph2TextRuns = paragraph2Elements.Where(e => e.Name.Equals(WordMl.TextRunName)).ToList();
            Assert.AreEqual(2, paragraph1TextRuns.Count);
            Assert.AreEqual(2, paragraph2TextRuns.Count);
            Assert.AreEqual(this.dateValues[0], paragraph1TextRuns[0].Value);
            Assert.AreEqual(this.dateValues[1], paragraph2TextRuns[0].Value);
            Assert.AreEqual("1", paragraph1TextRuns[1].Value);
            Assert.AreEqual("2", paragraph2TextRuns[1].Value);

            var thirdLevelStaticText1 = paragraph1Elements[2];
            var thirdLevelStaticText2 = paragraph2Elements[2];
            Assert.IsNotNull(thirdLevelStaticText1);
            Assert.IsNotNull(thirdLevelStaticText2);
            Assert.AreEqual(this.thirdLevelStaticText, thirdLevelStaticText1.Value);
            Assert.AreEqual(this.thirdLevelStaticText, thirdLevelStaticText2.Value);
            Assert.AreEqual(this.staticText, thirdLevelStaticText1.Name);
            Assert.AreEqual(this.staticText, thirdLevelStaticText2.Name);

            var firstLevelStaticElements = container.Elements(this.staticText).ToList();
            Assert.AreEqual(2, firstLevelStaticElements.Count);

            var firstLevelStaticText1 = firstLevelStaticElements[0];
            var firstLevelStaticText2 = firstLevelStaticElements[1];

            Assert.IsNotNull(firstLevelStaticText1);
            Assert.IsNotNull(firstLevelStaticText2);
            Assert.AreEqual(this.firstLevelStaticText, firstLevelStaticText1.Value);
            Assert.AreEqual(this.firstLevelStaticText, firstLevelStaticText2.Value);
            Assert.AreEqual(this.staticText, firstLevelStaticText1.Name);
            Assert.AreEqual(this.staticText, firstLevelStaticText2.Name);
        }
}
}
