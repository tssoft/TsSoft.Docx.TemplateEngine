using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Xml.Serialization;
using TsSoft.Docx.TemplateEngine.Tags;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    [TestClass]
    public class DataReaderTest
    {
        [TestMethod]
        public void TestReadText()
        {
            const string message = "Hello, world!";
            var testData = new TestData { Message = message };
            var dataReader = new DataReader<TestData>(testData);
            Assert.AreEqual(message, dataReader.ReadText("//Message"));
        }
    }

    [XmlRootAttribute]
    public class TestData
    {
        [XmlElement]
        public string Message { get; set; }
    }
}