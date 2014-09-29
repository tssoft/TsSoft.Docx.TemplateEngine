using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class DataReader<E>
    {
        private XDocument dataDocument;

        public DataReader(E dataEntity)
        {
            dataDocument = new XDocument();
            using (var writer = dataDocument.CreateWriter())
            {
                var serializer = new XmlSerializer(typeof(E));
                serializer.Serialize(writer, dataEntity);
            }
        }

        public string ReadText(string expression)
        {
            var textElement = dataDocument.XPathSelectElement(expression);
            return textElement.Value;
        }
    }
}