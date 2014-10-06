using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class TextParser : GeneralParser
    {
        public override void Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            this.ValidateStartTag(startElement, "Text");
            var tag = new TextTag { Expression = startElement.Value, TagNode = startElement };
            var processor = new TextProcessor { TextTag = tag };
            parentProcessor.AddProcessor(processor);
        }
    }
}