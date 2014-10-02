using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class TextParser : GeneralParser
    {
        public override ITagProcessor Parse(XElement startElement)
        {
            ValidateStartTag(startElement, "Text");
            var tag = new TextTag { Expression = startElement.Value };
            return new TextProcessor(tag);
        }
    }
}