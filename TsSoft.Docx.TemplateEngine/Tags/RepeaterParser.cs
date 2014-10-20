using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System;

    internal class RepeaterParser : GeneralParser
    {
        private const string TagName = "Repeater";
        private const string EndTagName = "EndRepeater";
        private const string ItemsTagName = "Items";
        private const string StartContentTagName = "Content";
        private const string EndContentTagName = "EndContent";
        private const string IndexTag = "ItemIndex";
        private const string ItemTag = "ItemText";

        private static Func<XElement, RepeaterElement> MakeElementCallback = element =>
            {
                var repeaterElement = new RepeaterElement
                {
                    Elements = element.Elements().Select(MakeElementCallback),
                    IsIndex = element.IsTag(IndexTag),
                    IsItem = element.IsTag(ItemTag),
                    XElement = element
                };
                if (repeaterElement.IsItem)
                {
                    repeaterElement.Expression = element.GetExpression();
                }
                return repeaterElement;
            };

        public override XElement Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            this.ValidateStartTag(startElement, TagName);
            var endRepeater = TryGetRequiredTag(startElement, EndTagName);
            //var itemsTag = TryGetRequiredTag(startElement, endRepeater, ItemsTagName);
            var itemsSource = startElement.GetExpression();

            if (string.IsNullOrEmpty(itemsSource))
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, "Items"));
            }

            var startContent = TryGetRequiredTag(startElement, endRepeater, StartContentTagName);
            var endContent = TryGetRequiredTag(startElement, endRepeater, EndContentTagName);

            IEnumerable<XElement> elementsBetween = TraverseUtils.ElementsBetween(startElement, endRepeater).ToList();
            var repeaterTag = new RepeaterTag
                {
                    Source = itemsSource,
                    StartContent = startContent,
                    EndContent = endContent,
                    StartRepeater = startElement,
                    EndRepeater = endRepeater,
                    MakeElementCallback = MakeElementCallback
                };

            var repeaterProcessor = new RepeaterProcessor
                {
                    RepeaterTag = repeaterTag,
                };

            if (elementsBetween.Any())
            {
                this.GoDeeper(repeaterProcessor, elementsBetween.First());
            }

            parentProcessor.AddProcessor(repeaterProcessor);

            return endRepeater;
        }

        private bool GoDeeper(ITagProcessor parentProcessor, XElement element)
        {
            var endReached = false;
            do
            {
                if (element.IsSdt())
                {
                    switch (this.GetTagName(element).ToLower())
                    {
                        case "endcontent":
                            return true;
                        case "itemtext":
                        case "itemindex":
                            break;
                        default:
                            element = this.ParseSdt(parentProcessor, element);
                            break;
                    }
                }
                else if (element.HasElements)
                {
                    endReached = this.GoDeeper(parentProcessor, element.Elements().First());
                }
                element = element.NextElement();
            }
            while (element != null && !endReached);
            return endReached;
        }
    }
}
