using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class TextParser : GeneralParser
    {
        public override void Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            ValidateStartTag(startElement, "Text");
            var tag = new TextTag { Expression = startElement.Value };
            var processor = new TextProcessor { TextTag = tag };
            parentProcessor.AddProcessor(processor);
            base.Parse(parentProcessor, startElement.NextElement());
        }
    }
}