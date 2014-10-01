using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags;

namespace TsSoft.Docx.TemplateEngine.Parsers
{
    internal class TextParser : GeneralParser
    {
        public TextTag Do(XElement startElement)
        {
            ValidateStartTag(startElement, "Text");
            return new TextTag { Expression = startElement.Value };
        }
    }
}