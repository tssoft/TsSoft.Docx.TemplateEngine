using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class TextTag
    {
        /// <summary>
        /// XPath expression
        /// </summary>
        public string Expression { get; set; }

        public XElement TagNode { get; set; }
    }
}