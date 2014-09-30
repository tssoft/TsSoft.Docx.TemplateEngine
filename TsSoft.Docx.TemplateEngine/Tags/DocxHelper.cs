using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class DocxHelper
    {
        public static readonly XNamespace WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static XElement CreateTextElement(string text)
        {
            return new XElement(WordMlNamespace + "r", new XElement(WordMlNamespace + "t", text));
        }
    }
}