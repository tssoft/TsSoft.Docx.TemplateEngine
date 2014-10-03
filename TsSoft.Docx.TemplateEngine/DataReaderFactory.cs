using System;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace TsSoft.Docx.TemplateEngine
{
    internal static class DataReaderFactory
    {
        public static DataReader CreateReader<E>(E dataEntity)
        {
            if (dataEntity == null)
            {
                throw new ArgumentNullException();
            }

            var dataDocument = new XDocument();
            using (var writer = dataDocument.CreateWriter())
            {
                var serializer = new XmlSerializer(typeof(E));
                serializer.Serialize(writer, dataEntity);
            }
            return CreateReader(dataDocument);
        }

        public static DataReader CreateReader(string dataXml)
        {
            return CreateReader(XDocument.Parse(dataXml));
        }

        public static DataReader CreateReader(XDocument dataDocument)
        {
            return new DataReader(dataDocument.Root);
        }
    }
}