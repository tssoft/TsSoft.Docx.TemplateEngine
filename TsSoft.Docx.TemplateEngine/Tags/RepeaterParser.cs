using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class RepeaterParser : GeneralParser
    {
        private const string TagName = "Repeater";
        private const string EndTagName = "EndRepeater";
        private const string ItemsTagName = "Items";
        private const string StartContentTagName = "Content";
        private const string EndContentTagName = "EndContent";
        private const string IndexTag = "ItemIndex";
        private const string ItemTag = "Item";

        public override void Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            ValidateStartTag(startElement, TagName);
            var endRepeater = TryGetRequiredTag(startElement, EndTagName);
            var itemsTag = TryGetRequiredTag(startElement, endRepeater, ItemsTagName);
            var startContent = TryGetRequiredTag(startElement, endRepeater, StartContentTagName);
            var endContent = TryGetRequiredTag(startElement, endRepeater, EndContentTagName);


            IEnumerable<RepeaterElement> repeaterElements = TraverseUtils.ElementsBetween(startContent, endContent).Select(makeRepeaterElement);
            var repeaterTag = new RepeaterTag
                {
                    Source = itemsTag.Value,
                    StartContent = startContent,
                    EndContent = endContent,
                    Content = repeaterElements,
                    StartRepeater = startElement,
                    EndRepeater = endRepeater,
                };

            var repeaterProcessor = new RepeaterProcessor
                {
                    RepeaterTag = repeaterTag,
                };


            //TODO implement
            var foundTags = new List<XElement>();
            foreach (var foundTag in foundTags)
            {
                base.Parse(parentProcessor, foundTag);
            }

            parentProcessor.AddProcessor(repeaterProcessor);
        }

        private RepeaterElement makeRepeaterElement(XElement xElement)
        {
            var repeaterElement = new RepeaterElement
            {
                Elements = xElement.Elements().Select(makeRepeaterElement),
                IsIndex = xElement.IsTag(IndexTag),
                IsItem = xElement.IsTag(ItemTag),
                XElement = xElement
            };
            if (repeaterElement.IsItem)
            {
                repeaterElement.Expression = xElement.GetExpression();
            }
            return repeaterElement;
        }
    }
}
