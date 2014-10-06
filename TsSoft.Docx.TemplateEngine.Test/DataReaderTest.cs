using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;
using TsSoft.Commons.Utils;

namespace TsSoft.Docx.TemplateEngine.Test
{
    [TestClass]
    public class DataReaderTest
    {
        [TestMethod]
        public void TestReadText()
        {
            const string Message = "Hello, world!";
            var testData = new DataReaderTestData { Message = Message };
            var dataReader = DataReaderFactory.CreateReader(testData);
            Assert.AreEqual(Message, dataReader.ReadText("//Message"));
        }

        [TestMethod]
        public void TestGetReader()
        {
            var document = this.GetXmlDocument();
            var dataReader = DataReaderFactory.CreateReader(document);

            const string Path = "/Test/Certificates/Certificate";
            var reader = dataReader.GetReader(Path);
            Assert.IsNotNull(reader);

            var node = document.SelectSingleNode(Path);
            var expectedReader = DataReaderFactory.CreateReader(node);
            Assert.AreEqual(expectedReader, reader);

            const string WrongPath = "/Test/Documents/Document";
            reader = dataReader.GetReader(WrongPath);
            Assert.IsNull(reader);
        }

        [TestMethod]
        public void TestGetReaders()
        {
            var document = GetXmlDocument();
            var dataReader = DataReaderFactory.CreateReader(document);

            const string Path = "/Test/Certificates/Certificate";
            var readers = dataReader.GetReaders(Path);

            Assert.IsNotNull(readers);

            var nodes = document.SelectNodes(Path);
            Assert.AreEqual(nodes.Count, readers.Count());

            var dataReadersEnumerator = readers.GetEnumerator();
            foreach (XmlNode node in nodes)
            {
                var expectedReader = DataReaderFactory.CreateReader(node);
                dataReadersEnumerator.MoveNext();
                Assert.AreEqual(expectedReader, dataReadersEnumerator.Current);
            }

            const string WrongPath = "/Test/Documents/Document";
            readers = dataReader.GetReaders(WrongPath);

            Assert.IsNotNull(readers);
            Assert.IsFalse(readers.Any());
        }

        [TestMethod]
        [ExpectedException(typeof(System.ArgumentNullException))]
        public void TestGetReaderNullPath()
        {
            var dataReader = DataReaderFactory.CreateReader(this.GetXmlDocument());
            dataReader.GetReader(null);
        }

        [TestMethod]
        [ExpectedException(typeof(System.ArgumentNullException))]
        public void TestGetReadersNullPath()
        {
            var dataReader = DataReaderFactory.CreateReader(this.GetXmlDocument());
            var readers = dataReader.GetReaders(null);
            readers.GetEnumerator().MoveNext();
        }

        private XmlDocument GetXmlDocument()
        {
            var stream = AssemblyResourceHelper.GetResourceStream(this, "DataReaderTest.xml");
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(stream);
            return xmlDoc;
        }
    }

    [XmlRoot]
    public class DataReaderTestData
    {
        [XmlElement]
        public string Message { get; set; }
    }
}