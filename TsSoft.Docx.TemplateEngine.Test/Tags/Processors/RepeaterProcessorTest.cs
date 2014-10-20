namespace TsSoft.Docx.TemplateEngine.Test.Tags.Processors
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Moq;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml.Linq;
    using TsSoft.Docx.TemplateEngine.Tags;
    using TsSoft.Docx.TemplateEngine.Tags.Processors;

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

        [TestMethod]
        public void TestProcess()
        {
            var processor = new RepeaterProcessor();

            var startRepeater = new XElement(WordMl.SdtName, "//test/certificates");
            var endRepeater = new XElement(WordMl.SdtName, "EndRepeater");
            //var endContent = new XElement(WordMl.SdtName, "EndContent");
            //var startContent = new XElement(WordMl.SdtName, "StartContent");
           // var items = new XElement(WordMl.SdtName, "Items");
            
            const string Index = "index";
            var indexElement = new XElement(Index);

            const string ThirdLevelStaticText = "Very static. Wow. Much text.";
            const string StaticText = "statictext";
            var thirdLevelStaticElement = new XElement(StaticText, ThirdLevelStaticText);

            const string FirstLevelStaticText = "This text must render as it is";
            var firstLevelStaticElement = new XElement(StaticText, FirstLevelStaticText);

            const string SecondLevelStaticText = "Just as its brother above, this text must render as it is, too";
            var secondLevelStaticElement = new XElement(StaticText, SecondLevelStaticText);

            var secondLevelElement = new XElement(WordMl.ParagraphName);
            var itemElement = new XElement("item");
            itemElement.Value = "./Date";
            secondLevelElement.Add(itemElement);
            secondLevelElement.Add(indexElement);
            secondLevelElement.Add(thirdLevelStaticElement);

            var firstLevelItem = new XElement("item");
            firstLevelItem.Value = "./Subject";

            const string Wrapper = "wrapper";
            var firstLevelElement = new XElement("wrapper");
            firstLevelElement.Add(secondLevelElement);
            firstLevelElement.Add(secondLevelStaticElement);

            var body = new XElement("body", startRepeater, firstLevelItem, firstLevelElement, firstLevelStaticElement, endRepeater);

            Console.WriteLine(body.ToString());

            const string Subject1Value = "Subject1";
            var subject1 = new XElement("subject") { Value = Subject1Value };
            const string Date1Value = "10.01.2014";
            var date1 = new XElement("date") { Value = Date1Value };

            const string Subject2Value = "Subject2";
            var subject2 = new XElement("subject") { Value = Subject2Value };
            const string Date2Value = "22.02.2014";
            var date2 = new XElement("date") { Value = Date2Value };

            var certificate1 = new XElement("certificate", subject1, date1);
            var certificate2 = new XElement("certificate", subject2, date2);

            var dataReaderMock = new Mock<DataReader>();

            const string XPath = "//test/certificates";
            dataReaderMock.Setup(d => d.GetReaders(XPath))
                .Returns(() => new List<DataReader>
                {
                    new DataReader(certificate1), 
                    new DataReader(certificate2)
                });

            processor.DataReader = dataReaderMock.Object;

            processor.RepeaterTag = new RepeaterTag
                {
                    //EndContent = endContent,
                    //StartContent = startContent,
                    Source = XPath,
                    StartRepeater = startRepeater,
                    EndRepeater = endRepeater,
                    MakeElementCallback = MakeElementCallback
                };

            processor.Process();

            Console.WriteLine(body.ToString());

            this.ValidateTagsRemoved(body);

            var subjects = body.Elements(WordMl.ParagraphName).ToList();
            Assert.AreEqual(6, body.Elements().Count());

            Assert.AreEqual(2, subjects.Count);
            Assert.AreEqual(Subject1Value, subjects[0].Value);
            Assert.AreEqual(Subject2Value, subjects[1].Value);

            var wrappers = body.Elements(Wrapper).ToList();
            Assert.AreEqual(2, wrappers.Count);

            var secondLevelElements1 = wrappers[0].Elements().ToList();
            var secondLevelElements2 = wrappers[1].Elements().ToList();

            Assert.AreEqual(2, secondLevelElements1.Count);
            Assert.AreEqual(2, secondLevelElements2.Count);

            var secondLevelStaticText1 = secondLevelElements1[1];
            var secondLevelStaticText2 = secondLevelElements1[1];
            Assert.IsNotNull(secondLevelStaticText1);
            Assert.IsNotNull(secondLevelStaticText2);
            Assert.AreEqual(SecondLevelStaticText, secondLevelStaticText1.Value);
            Assert.AreEqual(SecondLevelStaticText, secondLevelStaticText2.Value);
            Assert.AreEqual(StaticText, secondLevelStaticText1.Name);
            Assert.AreEqual(StaticText, secondLevelStaticText2.Name);

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
            Assert.AreEqual(Date1Value, paragraph1TextRuns[0].Value);
            Assert.AreEqual(Date2Value, paragraph2TextRuns[0].Value);
            Assert.AreEqual("1", paragraph1TextRuns[1].Value);
            Assert.AreEqual("2", paragraph2TextRuns[1].Value);

            var thirdLevelStaticText1 = paragraph1Elements[2];
            var thirdLevelStaticText2 = paragraph2Elements[2];
            Assert.IsNotNull(thirdLevelStaticText1);
            Assert.IsNotNull(thirdLevelStaticText2);
            Assert.AreEqual(ThirdLevelStaticText, thirdLevelStaticText1.Value);
            Assert.AreEqual(ThirdLevelStaticText, thirdLevelStaticText2.Value);
            Assert.AreEqual(StaticText, thirdLevelStaticText1.Name);
            Assert.AreEqual(StaticText, thirdLevelStaticText2.Name);

            var firstLevelStaticElements = body.Elements(StaticText).ToList();
            Assert.AreEqual(2, firstLevelStaticElements.Count);

            var firstLevelStaticText1 = firstLevelStaticElements[0];
            var firstLevelStaticText2 = firstLevelStaticElements[1];

            Assert.IsNotNull(firstLevelStaticText1);
            Assert.IsNotNull(firstLevelStaticText2);
            Assert.AreEqual(FirstLevelStaticText, firstLevelStaticText1.Value);
            Assert.AreEqual(FirstLevelStaticText, firstLevelStaticText2.Value);
            Assert.AreEqual(StaticText, firstLevelStaticText1.Name);
            Assert.AreEqual(StaticText, firstLevelStaticText2.Name);
        }
    }
}
