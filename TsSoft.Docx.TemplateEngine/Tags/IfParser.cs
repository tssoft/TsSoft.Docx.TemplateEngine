using System.Collections.Generic;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System;
    using System.Linq;
    using System.Xml.Linq;

    using TsSoft.Docx.TemplateEngine.Tags.Processors;

    internal class IfParser : GeneralParser
    {
        private const string StartTagName = "If";
        private const string EndTagName = "EndIf";

        public override XElement Parse(ITagProcessor parentProcessor, XElement startElement)
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

            if (content.Any())
            {
                this.GoDeeper(ifProcessor, content.First());
            }

            parentProcessor.AddProcessor(ifProcessor);

            return endTag;
        }

        private void GoDeeper(ITagProcessor parentProcessor, XElement element)
        {
            do
            {
                if (element.IsSdt())
                {
                    element = this.ParseSdt(parentProcessor, element);
                }
                else if (element.HasElements)
                {
                    this.GoDeeper(parentProcessor, element.Elements().First());
                }
                element = element.NextElement();
            }
            while (element != null && (!element.IsSdt() || GetTagName(element).ToLower() != "endif"));
        }
    }
}
