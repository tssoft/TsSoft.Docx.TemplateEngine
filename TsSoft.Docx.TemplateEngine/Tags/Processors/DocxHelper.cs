using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    internal class DocxHelper
    {
        public static XElement CreateTextElement(string text)
        {
            return new XElement(WordMl.WordMlNamespace + "r", new XElement(WordMl.WordMlNamespace + "t", text));
        }
    }
}