using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags;

namespace TsSoft.Docx.TemplateEngine.Parsers
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

        public RepeaterTag Do(XElement openingElement)
        {
            ValidateStartTag(openingElement, TagName);
            var enclosingRepeaterTag = TryGetRequiredTag(openingElement, EndTagName);
            var itemsTag = TryGetRequiredTag(openingElement, enclosingRepeaterTag, ItemsTagName);
            var startContent = TryGetRequiredTag(openingElement, enclosingRepeaterTag, StartContentTagName);
            var endContent = TryGetRequiredTag(openingElement, enclosingRepeaterTag, EndContentTagName);

            IEnumerable<RepeaterElement> repeaterElements = TraverseUtils.ElementsBetween(startContent, endContent).Select(makeRepeaterElement);
            return new RepeaterTag
            {
                Source = itemsTag.Value,
                Start = startContent,
                End = endContent,
                Content = repeaterElements
            };
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
