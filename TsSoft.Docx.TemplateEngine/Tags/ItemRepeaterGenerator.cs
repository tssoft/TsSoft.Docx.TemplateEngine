using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class ItemRepeaterGenerator
    {
        private const string ItemRepeaterTagName = "itemrepeater";
        private const string EndItemRepeaterTagName = "enditemrepeater";

        private const string IndexTag = "RItemIndex";
        private const string ItemTag = "RItemText";
        private const string ItemIf = "RItemIf";
        private const string EndItemIf = "REndIf";
        private const string ItemHtmlContentTag = "RItemHtmlContent";
       
        private static Func<XElement, ItemRepeaterElement> MakeElementCallback = element =>
        {
            var itemRepeaterElement = new ItemRepeaterElement()
            {
                Elements = element.Elements().Select(MakeElementCallback),
                IsIndex = element.IsTag(IndexTag),
                IsItem = element.IsTag(ItemTag),
                IsItemIf = element.IsTag(ItemIf),
                IsEndItemIf = element.IsTag(EndItemIf),
                IsItemRepeater = element.IsTag(ItemRepeaterTagName),
                IsEndItemRepeater = element.IsTag(EndItemRepeaterTagName),
                IsItemHtmlContent = element.IsTag(ItemHtmlContentTag),
                XElement = element
            };
            if (itemRepeaterElement.IsItem || itemRepeaterElement.IsItemIf || itemRepeaterElement.IsItemRepeater)
            {
                itemRepeaterElement.Expression = element.GetExpression();
            }
            itemRepeaterElement.StartTag = itemRepeaterElement.XElement;
            if (itemRepeaterElement.IsItemRepeater || itemRepeaterElement.IsItemIf)
            {                
                itemRepeaterElement.EndTag = FindEndTag(itemRepeaterElement.StartTag,
                    (itemRepeaterElement.IsItemRepeater) ? ItemRepeaterTagName : ItemIf,
                    (itemRepeaterElement.IsItemRepeater) ? EndItemRepeaterTagName : EndItemIf
                );
            }
            return itemRepeaterElement;
        };        

        public static bool IsItemRepeaterElement(XElement element)
        {
            return element.IsSdt() &&
                   (element.IsTag(IndexTag) || element.IsTag(ItemHtmlContentTag) || element.IsTag(ItemTag) ||
                    element.IsTag(ItemIf) || element.IsTag(EndItemIf) || element.IsTag(EndItemRepeaterTagName));
        }

        public XElement Generate(ItemRepeaterTag tag, IEnumerable<DataReader> dataReaders, XElement previous = null, XElement parent = null, bool forcedElementSave = false)
        {
            var startElement = tag.StartItemRepeater;
            var endElement = tag.EndItemRepeater;
            var inlineWrapping = this.CheckInlineWrappingMode(startElement, endElement);
            var itemRepeaterElements =
                TraverseUtils.SecondElementsBetween(startElement, endElement).Select(MakeElementCallback).ToList();
            var flgCleanUpElements = previous == null;
            XElement current;
            if (inlineWrapping && flgCleanUpElements)
            {
                current = startElement;
            }
            else
            {
                if (previous == null)
                {
                    current = startElement.Parent.Name.Equals(WordMl.ParagraphName) ? startElement.Parent : startElement;
                }
                else
                {
                    current = previous;                    
                }
            }
            for (var index = 1; index <= dataReaders.Count(); index++)
            {
                XElement nestedRepeaterEndElementTmp = null;
                XElement endIfElementTmp = null;
                current = this.ProcessElements(
                                               itemRepeaterElements,
                                               dataReaders.ElementAt(index - 1),
                                               current,
                                               (inlineWrapping && (parent != null)) ? parent : null,
                                               index,
                                               ref nestedRepeaterEndElementTmp,
                                               ref endIfElementTmp,                                               
                                               inlineWrapping && (parent != null)
                );
            }
            if (flgCleanUpElements && !forcedElementSave)
            {
                foreach (var itemRepeaterElement in itemRepeaterElements)
                {
                    itemRepeaterElement.XElement.Remove();
                }                                
                startElement.Remove();                
                endElement.Remove();
                
            }            
            return current;
        }       

        private bool CheckInlineWrappingMode(XElement startElement, XElement endElement)
        {
            return startElement.Parent.Equals(endElement.Parent);
        }

        private void ProcessIfElement(ItemRepeaterElement ifElement, DataReader dataReader, ref XElement endIfElement)
        {
            const string IsNotLastElementCondition = "isnotlastelement";
            const string IsLastElementCondition = "islastelement";
            
            var expression = ifElement.Expression;
            if (expression.Equals(IsNotLastElementCondition) || expression.Equals(IsLastElementCondition))
            {
                var cond = expression.Equals(IsNotLastElementCondition) ^ dataReader.IsLast;
                if (!cond)
                {
                    endIfElement = ifElement.EndTag;
                }
            }
            else
            {
                var condition = bool.Parse(dataReader.ReadText(expression));
                if (!condition)
                {
                    endIfElement = ifElement.EndTag;
                }                
            }                                
        }


        private bool CheckNestedElementForContinue(ItemRepeaterElement firstItemRepeaterElement,
                                                   ItemRepeaterElement currentItemRepeaterElement,
                                                   XElement nestedRepeaterEndElement)
        {
            if (nestedRepeaterEndElement != null)
            {
                return
                    TraverseUtils.SecondElementsBetween(firstItemRepeaterElement.XElement, nestedRepeaterEndElement)
                                 .Contains(currentItemRepeaterElement.XElement);
            }
            return false;
        }

        private bool CheckAndProcessEndItemRepeaterElementForContinue(ItemRepeaterElement currentItemRepeaterElement,
                                                            ref XElement nestedRepeaterEndElement)
        {
            if (currentItemRepeaterElement.IsEndItemRepeater &&
                currentItemRepeaterElement.XElement.Equals(nestedRepeaterEndElement))
            {
                nestedRepeaterEndElement = null;
                return true;                
            }
            return false;
        }

        private bool CheckAndProcessStartIfElementForContinue(ItemRepeaterElement currentItemRepeaterElement,
                                                              DataReader dataReader,
                                                              ref XElement endIfElement)
        {
            if (currentItemRepeaterElement.IsItemIf)
            {
                this.ProcessIfElement(currentItemRepeaterElement, dataReader, ref endIfElement);
                return true;
            }
            return false;
        }

        private bool CheckNestedConditionElementForContinue(ItemRepeaterElement currentItemRepeaterElement,
                                                            XElement endIfElement)
        {
            return (endIfElement != null) && !currentItemRepeaterElement.XElement.Name.Equals(WordMl.ParagraphName);
        }

        private bool CheckAndProcessEndIfElementForContinue(ItemRepeaterElement currentItemRepeaterElement,
                                                  ref XElement endIfElement)
        {
            if (currentItemRepeaterElement.IsEndItemIf && currentItemRepeaterElement.XElement.Equals(endIfElement))
            {
                endIfElement = null;
                return true;
            }
            return false;
        }
        
        private XElement ProcessElements(IEnumerable<ItemRepeaterElement> elements, DataReader dataReader, XElement start, XElement parent, int index, ref XElement nestedRepeaterEndElement, ref XElement endIfElement, bool nestedElement = false)
        {
            XElement result = null;
            XElement previous = start;

            foreach (var itemRepeaterElement in elements)
            {               
                var flgStucturedElementProcessed = this.CheckAndProcessStartIfElementForContinue(itemRepeaterElement,
                                                                                                  dataReader,
                                                                                                  ref endIfElement)
                                                    ||
                                                    this.CheckAndProcessEndIfElementForContinue(itemRepeaterElement,
                                                                                                ref endIfElement)
                                                    ||
                                                    this.CheckAndProcessEndItemRepeaterElementForContinue(
                                                        itemRepeaterElement, ref nestedRepeaterEndElement);                                                                                
                if (!flgStucturedElementProcessed)
                {                    
                    var flgNestedElementCheckedForContinue = this.CheckNestedConditionElementForContinue(
                                                              itemRepeaterElement, endIfElement)
                                                              ||
                                                              this.CheckNestedElementForContinue(elements.First(),
                                                                                                 itemRepeaterElement,
                                                                                                 nestedRepeaterEndElement);
                    if (!flgNestedElementCheckedForContinue)
                    {
                        if (itemRepeaterElement.IsItemHtmlContent)
                        {
                            result = HtmlContentProcessor.MakeHtmlContentProcessed(itemRepeaterElement.XElement,
                                                                                   dataReader.ReadText(
                                                                                       itemRepeaterElement.Expression),
                                                                                   true);
                        }
                        else if (itemRepeaterElement.IsItemRepeater)
                        {
                            var itemRepeaterTag = new ItemRepeaterTag()
                                {
                                    StartItemRepeater = itemRepeaterElement.StartTag,
                                    EndItemRepeater = itemRepeaterElement.EndTag,
                                    Source = itemRepeaterElement.Expression
                                };
                            var itemRepeaterGenerator = new ItemRepeaterGenerator();
                            previous = itemRepeaterGenerator.Generate(itemRepeaterTag,
                                                                      dataReader.GetReaders(itemRepeaterTag.Source),
                                                                      previous, parent);
                            nestedRepeaterEndElement = itemRepeaterTag.EndItemRepeater;
                            result = null;
                        }
                        else if (itemRepeaterElement.IsIndex)
                        {
                            result = DocxHelper.CreateTextElement(itemRepeaterElement.XElement,
                                                                  itemRepeaterElement.XElement.Parent,
                                                                  index.ToString(CultureInfo.CurrentCulture),
                                                                  !nestedElement);
                        }
                        else if (itemRepeaterElement.IsItem)
                        {
                            result = DocxHelper.CreateTextElement(itemRepeaterElement.XElement,
                                                                  itemRepeaterElement.XElement.Parent,
                                                                  dataReader.ReadText(itemRepeaterElement.Expression),
                                                                  !nestedElement);
                        }
                        else
                        {
                            var element = new XElement(itemRepeaterElement.XElement);
                            element.RemoveNodes();
                            result = element;
                            if (itemRepeaterElement.HasElements)
                            {
                                var parsedLastElement = this.ProcessElements(itemRepeaterElement.Elements, dataReader,
                                                                             previous,
                                                                             result, index, ref nestedRepeaterEndElement,
                                                                             ref endIfElement, true);
                                if (itemRepeaterElement.Elements.Any(ire => ire.XElement.IsSdt()) && DocxHelper.IsEmptyParagraph(result))
                                {
                                    result = null;
                                }
                                if (
                                    itemRepeaterElement.Elements.Any(
                                        ire =>
                                        ire.IsItemRepeater && !this.CheckInlineWrappingMode(ire.StartTag, ire.EndTag)))
                                {
                                    previous = parsedLastElement;
                                }
                                
                            }
                            else
                            {
                                element.Value = itemRepeaterElement.XElement.Value;
                            }
                        }
                        if (result != null)
                        {
                            if (!nestedElement)
                            {
                                previous.AddAfterSelf(result);
                                previous = result;
                            }
                            else
                            {
                                parent.Add(result);
                            }
                        }
                    }
                    else
                    {
                        result = null;
                    }
                }
                else
                {
                    result = null;
                }
            }            
            return result ?? previous;
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
    }
}
