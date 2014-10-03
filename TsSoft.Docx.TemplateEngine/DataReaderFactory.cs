using System;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace TsSoft.Docx.TemplateEngine
{
    internal class DataReaderFactory : IDataReaderFactory
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

        DataReader IDataReaderFactory.CreateReader(string dataXml)
        {
            return CreateReader(dataXml);
        }

        DataReader IDataReaderFactory.CreateReader(XDocument dataDocument)
        {
            return CreateReader(dataDocument);
        }

        DataReader IDataReaderFactory.CreateReader<E>(E dataEntity)
        {
            return CreateReader(dataEntity);
        }
    }
}