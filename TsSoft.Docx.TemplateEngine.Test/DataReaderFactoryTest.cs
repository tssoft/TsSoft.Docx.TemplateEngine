using System;

namespace TsSoft.Docx.TemplateEngine.Test
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Diagnostics.CodeAnalysis;
    using System.Xml.Linq;
    using System.Xml.Serialization;

    [TestClass]
    public class DataReaderFactoryTest
    {
        [TestMethod]
        public void TestCreateReader()
        {
            var testData = new DataReaderFactoryTestData
            {
                Message = "test factory",
            };

            var element = new XElement("testdata");
            element.Add(new XElement("message", testData.Message));
            var expected = new DataReader(element);
            var actual = DataReaderFactory.CreateReader(testData);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        [ExpectedException(typeof(System.ArgumentNullException))]
        public void TestCreateReaderNullArgument()
        {
            DataReaderFactory.CreateReader<DataReaderTestData>(null);
        }

        [TestMethod]
        public void TestCreateReaderStringXml()
        {
            const string xml = "<entities>test</entities>";
            var reader = DataReaderFactory.CreateReader(xml);
            

        }
    }

    public class CustomAttribute : Attribute
    {
        public int SomeValue { get; set; }

        public CustomAttribute(int someValue)
        {
            SomeValue = someValue;
        }
    }

    [SuppressMessage("StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Reviewed. Suppression is OK here."),
    XmlRoot(Namespace = "", ElementName = "TestData", IsNullable = false)]
    public class DataReaderFactoryTestData
    {
        [XmlElement]
        public string Message { get; set; }
    }
}