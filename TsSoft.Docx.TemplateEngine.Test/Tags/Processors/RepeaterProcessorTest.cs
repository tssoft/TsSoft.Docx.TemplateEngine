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
            const string IndexAndDate = "indexAndDate";
            var indexAndDateElement = new XElement(IndexAndDate, dateElement, indexElement);

            const string Wrapper = "wrapper";
            var wrapperElement = new XElement(Wrapper, indexAndDateElement);

            var body = new XElement("body", startRepeater, items, startContent, subjectElement, wrapperElement, endContent, endRepeater);

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
                }
            };
            var secondLevelContent = new List<RepeaterElement>
            {
                new RepeaterElement
                {
                    XElement = indexAndDateElement,
                    Elements = thirdLevelContent
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

            var subjects = body.Elements(WordMl.RName).ToList();
            Assert.AreEqual(4, body.Elements().Count());

            Assert.AreEqual(2, subjects.Count);
            Assert.AreEqual(Subject1Value, subjects[0].Value);
            Assert.AreEqual(Subject2Value, subjects[1].Value);

            var wrappers = body.Elements(Wrapper).ToList();
            Assert.AreEqual(2, wrappers.Count);

            var wrapper1IndexAndDate = wrappers[0].Elements(IndexAndDate).ToList();
            var wrapper2IndexAndDate = wrappers[1].Elements(IndexAndDate).ToList();
            Assert.AreEqual(1, wrapper1IndexAndDate.Count);
            Assert.AreEqual(1, wrapper2IndexAndDate.Count);

            var wrapper1Replaced = wrapper1IndexAndDate.Elements(WordMl.RName).ToList();
            var wrapper2Replaced = wrapper2IndexAndDate.Elements(WordMl.RName).ToList();
            Assert.AreEqual(2, wrapper1Replaced.Count);
            Assert.AreEqual(2, wrapper2Replaced.Count);
            Assert.AreEqual(Date1Value, wrapper1Replaced[0].Value);
            Assert.AreEqual(Date2Value, wrapper2Replaced[0].Value);
            Assert.AreEqual("1", wrapper1Replaced[1].Value);
            Assert.AreEqual("2", wrapper2Replaced[1].Value);
        }
    }
}
