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
            if (itemRepeaterElement.IsItemRepeater)
            {                
                itemRepeaterElement.EndTag = FindEndTag(itemRepeaterElement.StartTag, ItemRepeaterTagName,
                                                        EndItemRepeaterTagName);
            }
            return itemRepeaterElement;
        };        

        public XElement Generate(ItemRepeaterTag tag, IEnumerable<DataReader> dataReaders, XElement previous = null)
        {
            var startElement = tag.StartItemRepeater;
            var endElement = tag.EndItemRepeater;
            var itemRepeaterElements =
                TraverseUtils.ElementsBetween(startElement, endElement).Select(MakeElementCallback).ToList();
            var flgCleanUpElements = previous == null;
            XElement current;
            if (previous == null)
            {
                current = startElement.Parent.Name.Equals(WordMl.ParagraphName) ? startElement.Parent : startElement;
            }
            else
            {
                //var tmpDoc = new XDocument();
                //tmpDoc.Add(previous);
                current = previous;                
            }
            for (var index = 1; index <= dataReaders.Count(); index++)
            {
                XElement tmp = null;
                current = this.ProcessElements(itemRepeaterElements, dataReaders.ElementAt(index - 1), current, null,
                                               index, ref tmp);
            }
            if (flgCleanUpElements)
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

        private XElement ProcessElements(IEnumerable<ItemRepeaterElement> elements, DataReader dataReader, XElement start, XElement parent, int index, ref XElement nestedRepeaterEndElement, bool nestedParagraph = false)
        {
            XElement result = null;
            XElement previous = start;
            foreach (var itemRepeaterElement in elements)
            {              
                if (nestedRepeaterEndElement != null)
                {
                    if (TraverseUtils.ElementsBetween(elements.First().XElement, nestedRepeaterEndElement)
                                     .Contains(itemRepeaterElement.XElement))
                    {
                        continue;
                    }
                }
                if (itemRepeaterElement.IsEndItemRepeater && itemRepeaterElement.XElement.Equals(nestedRepeaterEndElement))
                {
                    nestedRepeaterEndElement = null;
                }
                else if (itemRepeaterElement.IsItemHtmlContent)
                {
                    result = HtmlContentProcessor.MakeHtmlContentProcessed(itemRepeaterElement.XElement,
                                                                           dataReader.ReadText(itemRepeaterElement.Expression),
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
                                                              previous ?? new XElement(parent));
                    nestedRepeaterEndElement = itemRepeaterTag.EndItemRepeater;
                    result = null;
                    continue;
                }               
                else if (itemRepeaterElement.IsIndex)
                {
                    result = DocxHelper.CreateTextElement(itemRepeaterElement.XElement,
                                                          itemRepeaterElement.XElement.Parent,
                                                          index.ToString(CultureInfo.CurrentCulture),
                                                          !nestedParagraph);
                }
                else if (itemRepeaterElement.IsItem)
                {
                    result = DocxHelper.CreateTextElement(itemRepeaterElement.XElement,
                                                          itemRepeaterElement.XElement.Parent,
                                                          dataReader.ReadText(itemRepeaterElement.Expression),
                                                          !nestedParagraph);
                }
                else
                {
                    var element = new XElement(itemRepeaterElement.XElement);
                    element.RemoveNodes();                    
                    result = element;
                    if (itemRepeaterElement.HasElements)
                    {
                        this.ProcessElements(itemRepeaterElement.Elements, dataReader, null, result, index, ref nestedRepeaterEndElement, true);
                    }
                    else
                    {
                        element.Value = itemRepeaterElement.XElement.Value;
                    }
                }
                if (result != null)
                {
                    if (previous != null)
                    {
                        if (false)
                        {                            
                            //previous.Elements().ToList().Add(result);
                            var nElement = new XElement(previous);
                            nElement.AddAfterSelf(previous);
                        }
                        else
                        {
                            previous.AddAfterSelf(result);
                            var x = new XElement(previous);                                                        
                        }
                        previous = result;
                    }
                    else
                    {
                        parent.Add(result);
                    }
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
