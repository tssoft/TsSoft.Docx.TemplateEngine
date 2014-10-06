using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class TextParser : GeneralParser
    {
        public override void Parse(ITagProcessor parentProcessor, XElement root)
        {
            this.ValidateStartTag(root, "Text");
            var tag = new TextTag { Expression = root.Value, TagNode = root };
            var processor = new TextProcessor { TextTag = tag };
            parentProcessor.AddProcessor(processor);
        }
    }
}