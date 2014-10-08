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
        [TestMethod]
        public void TestDo()
        {
            var processor = new RepeaterProcessor();

            var startRepeater = new XElement(WordMl.SdtName, "StartRepeater");
            var endRepeater = new XElement(WordMl.SdtName, "EndRepeater");
            var endContent = new XElement(WordMl.SdtName, "EndContent");
            var startContent = new XElement(WordMl.SdtName, "StartContent");
            var items = new XElement(WordMl.SdtName, "Items");

            const string Subject = "subject";
            var subjectElement = new XElement(Subject);

            const string Date = "date";
            var dateElement = new XElement(Date);
            const string Index = "index";
            var indexElement = new XElement(Index);

            const string ThirdLevelStaticText = "Very static. Wow. Much text.";
            const string StaticText = "StaticText";
            var thirdLevelStaticElement = new XElement(StaticText, ThirdLevelStaticText);

            var indexAndDateElement = new XElement(WordMl.ParagraphName, dateElement, indexElement, thirdLevelStaticElement);

            const string FirstLevelStaticText = "This text must render as it is";
            var firstLevelStaticElement = new XElement(StaticText, FirstLevelStaticText);
            
            const string SecondLevelStaticText = "Just as its brother above, this text must render as it is, too";
            var secondLevelStaticElement = new XElement(StaticText, SecondLevelStaticText);
            
            const string Wrapper = "wrapper";
            var wrapperElement = new XElement(Wrapper, indexAndDateElement, secondLevelStaticElement);

            var body = new XElement("body", startRepeater, items, startContent, subjectElement, wrapperElement, firstLevelStaticElement, endContent, endRepeater);

            var thirdLevelContent = new List<RepeaterElement>
            {
                new RepeaterElement
                {
                    XElement = dateElement, 
                    IsItem = true, 
                    Expression = "./Date"
                }, 
                new RepeaterElement
                {
                    XElement = indexElement, 
                    IsIndex = true
                },  new RepeaterElement
                {
                    XElement = thirdLevelStaticElement, 
                }
            };
            var secondLevelContent = new List<RepeaterElement>
            {
                new RepeaterElement
                {
                    XElement = indexAndDateElement, 
                    Elements = thirdLevelContent
                }, 
                new RepeaterElement
                {
                    XElement = secondLevelStaticElement, 
                }
            };
            var firstLevelContent = new List<RepeaterElement>
            {
                new RepeaterElement
                {
                    Expression = "./Subject", 
                    IsItem = true, 
                    XElement = subjectElement
                }, 
                new RepeaterElement
                {
                    XElement = wrapperElement, 
                    Elements = secondLevelContent
                }, 
                new RepeaterElement
                    {
                        XElement = firstLevelStaticElement
                    }
            };

            Console.WriteLine(body.ToString());

            const string Subject1Value = "Subject1";
            var subject1 = new XElement("Subject") { Value = Subject1Value };
            const string Date1Value = "10.01.2014";
            var date1 = new XElement("Date") { Value = Date1Value };

            const string Subject2Value = "Subject2";
            var subject2 = new XElement("Subject") { Value = Subject2Value };
            const string Date2Value = "22.02.2014";
            var date2 = new XElement("Date") { Value = Date2Value };

            var certificate1 = new XElement("Certificate", subject1, date1);
            var certificate2 = new XElement("Certificate", subject2, date2);

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
                    Content = firstLevelContent, 
                    EndContent = endContent, 
                    StartContent = startContent, 
                    Source = XPath, 
                    StartRepeater = startRepeater, 
                    EndRepeater = endRepeater
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
