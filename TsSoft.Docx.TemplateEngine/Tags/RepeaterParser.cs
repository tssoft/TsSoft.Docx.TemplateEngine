using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        private const string IndexTag = "ItemIndex";
        private const string ItemTag = "ItemText";
        private const string ItemIfTag = "ItemIf";
        private const string EndItemIfTag = "EndIf";
        private const string ItemRepeaterTag = "ItemRepeater";
        private const string EndItemRepeaterTag = "EndItemRepeater";
        private const string ItemHtmlContentTag = "ItemHtmlContent";
        private const string ItemTableTag = "ItemTable";
        private const string EndItemTableTag = "EndItemTable";



        public static Func<XElement, RepeaterElement> MakeElementCallback = element =>
            {
                var repeaterElement = new RepeaterElement
                {
                    Elements = element.Elements().Select(MakeElementCallback),
                    IsIndex = element.IsTag(IndexTag),
                    IsItem = element.IsTag(ItemTag),
                    IsItemIf = element.IsTag(ItemIfTag),                    
                    IsEndItemIf = element.IsTag(EndItemIfTag),
                    IsItemHtmlContent = element.IsTag(ItemHtmlContentTag),
                    IsItemRepeater = element.IsTag(ItemRepeaterTag),
                    IsItemTable = element.IsTag(ItemTableTag),
                    IsEndItemTable = element.IsTag(EndItemTableTag),
                    XElement = element,
                    StartTag = element                 
                };
                if (repeaterElement.IsItem || repeaterElement.IsItemHtmlContent || repeaterElement.IsItemRepeater || repeaterElement.IsItemIf || repeaterElement.IsItemTable)
                {
                    repeaterElement.Expression = element.GetExpression();
                }
                if (repeaterElement.IsItemRepeater)
                {                    
                    ParseStructuredElement(repeaterElement, ItemRepeaterTag, EndItemRepeaterTag);
                }
                else if (repeaterElement.IsItemTable)
                {
                    ParseStructuredElement(repeaterElement, ItemTableTag, EndItemTableTag);
                }
                else if (repeaterElement.IsItemIf)
                {
                    ParseStructuredElement(repeaterElement, ItemIfTag, EndItemIfTag);
                }
                return repeaterElement;
            };

        public override XElement Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            this.ValidateStartTag(startElement, TagName);
            var endRepeater = TryGetRequiredTag(startElement, EndTagName);
            var itemsSource = startElement.GetExpression(); 

            if (string.IsNullOrEmpty(itemsSource))
            {
                throw new Exception(MessageStrings.ItemsAreEmpty);
            }

            IEnumerable<XElement> elementsBetween = TraverseUtils.ElementsBetween(startElement, endRepeater).ToList();
            var repeaterTag = new RepeaterTag
                {
                    Source = itemsSource,
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

        private static void ParseStructuredElement(RepeaterElement repeaterElement, string startTag, string endTag)
        {
            repeaterElement.EndTag = FindEndTag(repeaterElement.StartTag, startTag, endTag);
            repeaterElement.TagElements = TraverseUtils.SecondElementsBetween(repeaterElement.StartTag,
                                                                              repeaterElement.EndTag)
                                                       .Select(MakeElementCallback);
        }

        private static XElement FindEndTag(XElement startTag, string startTagName, string endTagName)
        {
            var startTagsOpened = 1;
            var current = startTag;
            while (startTagsOpened > 0 && current != null)
            {
                var nextTagElements = TraverseUtils.NextTagElements(current, new Collection<string> { startTagName, endTagName }).ToList();
                var index = -1;
                while ((index < nextTagElements.Count) && (startTagsOpened != 0))
                {
                    index++;
                    if (nextTagElements[index].IsTag(startTagName))
                    {
                        startTagsOpened++;
                    }
                    else
                    {
                        startTagsOpened--;
                    }
                }
                current = nextTagElements[index];
            }
            if (current == null)
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, endTagName));
            }
            return current;
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
                        case "endrepeater":
                            return true;
                        case "itemrepeater":
                        case "itemtable":
                        case "itemtext":
                        case "itemif":
                        case "endif": 
                        case "enditemtable":
                        case "itemhtmlcontent":
                        case "itemindex":
                            break;
                        default:
                            if (!ItemRepeaterGenerator.IsItemRepeaterElement(element))
                            {
                                element = this.ParseSdt(parentProcessor, element);
                            }
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
