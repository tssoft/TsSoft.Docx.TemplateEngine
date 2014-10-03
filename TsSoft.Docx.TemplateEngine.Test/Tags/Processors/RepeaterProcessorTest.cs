using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags.Processors
{
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

            const string subject = "subject";
            var subjectElement = new XElement(subject);

            const string date = "date";
            var dateElement = new XElement(date);
            const string index = "index";
            var indexElement = new XElement(index);
            const string indexAndDate = "indexAndDate";
            var indexAndDateElement = new XElement(indexAndDate, dateElement, indexElement);


            const string wrapper = "wrapper";
            var wrapperElement = new XElement(wrapper, indexAndDateElement);

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

            const string subject1Value = "Subject1";
            var subject1 = new XElement("Subject") { Value = subject1Value };
            const string date1Value = "10.01.2014";
            var date1 = new XElement("Date") { Value = date1Value };

            const string subject2Value = "Subject2";
            var subject2 = new XElement("Subject") { Value = subject2Value };
            const string date2Value = "22.02.2014";
            var date2 = new XElement("Date") { Value = date2Value };


            var certificate1 = new XElement("Certificate", subject1, date1);
            var certificate2 = new XElement("Certificate", subject2, date2);

            var dataReaderMock = new Mock<DataReader>();

            const string xPath = "//test/certificates";
            dataReaderMock.Setup(d => d.GetReaders(xPath))
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
                    Source = xPath,
                    StartRepeater = startRepeater,
                    EndRepeater = endRepeater
                };
            
            processor.Process();


            Console.WriteLine(body.ToString());

            ValidateTagsRemoved(body);

            var subjects = body.Elements(WordMl.RName).ToList();
            Assert.AreEqual(4, body.Elements().Count());

            Assert.AreEqual(2, subjects.Count);
            Assert.AreEqual(subject1Value, subjects[0].Value);
            Assert.AreEqual(subject2Value, subjects[1].Value);

            var wrappers = body.Elements(wrapper).ToList();
            Assert.AreEqual(2, wrappers.Count);

            var wrapper1IndexAndDate = wrappers[0].Elements(indexAndDate).ToList();
            var wrapper2IndexAndDate = wrappers[1].Elements(indexAndDate).ToList();
            Assert.AreEqual(1, wrapper1IndexAndDate.Count);
            Assert.AreEqual(1, wrapper2IndexAndDate.Count);

            var wrapper1Replaced = wrapper1IndexAndDate.Elements(WordMl.RName).ToList();
            var wrapper2Replaced = wrapper2IndexAndDate.Elements(WordMl.RName).ToList();
            Assert.AreEqual(2, wrapper1Replaced.Count);
            Assert.AreEqual(2, wrapper2Replaced.Count);
            Assert.AreEqual(date1Value, wrapper1Replaced[0].Value);
            Assert.AreEqual(date2Value, wrapper2Replaced[0].Value);
            Assert.AreEqual("1", wrapper1Replaced[1].Value);
            Assert.AreEqual("2", wrapper2Replaced[1].Value);


        }
    }
}
