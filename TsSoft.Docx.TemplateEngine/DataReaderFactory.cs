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
            var dataReader = new DataReader(dataDocument.Root);
            return dataReader;
        }
    }
}