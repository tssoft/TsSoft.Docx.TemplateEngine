using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace TsSoft.Docx.TemplateEngine
{
    public class DocxGenerator<E>
    {
        public void GenerateDocx(Stream templateStream, Stream outputStream, E dataEntity)
        {
            using (var dataStream = new MemoryStream())
            using (var writer = new StreamWriter(dataStream))
            {
                var serializer = new XmlSerializer(typeof(E));
                serializer.Serialize(writer, dataEntity);
                dataStream.Seek(0, SeekOrigin.Begin);
                templateStream.Seek(0, SeekOrigin.Begin);
                templateStream.CopyTo(outputStream);
            }
        }

        internal void FillData(Stream outputStream, XmlDocument dataXmlDocument)
        { 
        }
    }
}