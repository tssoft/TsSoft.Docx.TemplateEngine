using System.Collections.Generic;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System;
    using System.Xml.Linq;

    using TsSoft.Docx.TemplateEngine.Tags.Processors;

    internal class IfParser : GeneralParser
    {
        private const string StartTagName = "If";
        private const string EndTagName = "EndIf";

        public override void Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            this.ValidateStartTag(startElement, StartTagName);
            var startTag = startElement;

            if (string.IsNullOrEmpty(startTag.Value))
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, "If"));
            }

            var endTag = TryGetRequiredTag(startElement, EndTagName);

            var content = TraverseUtils.ElementsBetween(startTag, endTag);
            var expression = startTag.GetExpression();

            var ifTag = new IfTag
                            {
                                Conidition = expression,
                                EndIf = endTag,
                                IfContent = content,
                                StartIf = startTag,
                            };

            var ifProcessor = new IfProcessor { Tag = ifTag };
            parentProcessor.AddProcessor(ifProcessor);

            this.GoDeeper(ifProcessor, ifTag.IfContent);

        }

        private void GoDeeper(ITagProcessor parentProcessor, IEnumerable<XElement> elements)
        {
            foreach (var element in elements)
            {
                if (element.IsSdt())
                {
                    this.ParseSdt(parentProcessor, element);
                }
                this.GoDeeper(parentProcessor, element.Elements());
            }
        }
    }
}
