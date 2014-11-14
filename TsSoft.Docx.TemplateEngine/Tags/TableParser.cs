using System;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;

    /// <summary>
    /// Parse table tag
    /// </summary>
    internal class TableParser : GeneralParser
    {
        private static Func<XElement, IEnumerable<TableElement>> MakeTableElementCallback = parentElement =>
            {
                ICollection<TableElement> tableElements = new Collection<TableElement>();
                var currentTagElement = TraverseUtils.NextTagElements(parentElement).FirstOrDefault();
                while (currentTagElement != null && currentTagElement.Ancestors().Any(element => element.Equals(parentElement)))
                {
                    var tableElement = ToTableElement(currentTagElement);
                    
                    if (tableElement.IsItem || tableElement.IsIndex || tableElement.IsItemIf || tableElement.IsItemRepeater)
                    {
                        tableElements.Add(tableElement);
                    }
                    if (tableElement.IsItemIf || tableElement.IsItemRepeater)
                    {
                        currentTagElement = tableElement.EndTag;
                    }
                    currentTagElement = TraverseUtils.NextTagElements(currentTagElement).FirstOrDefault();
                }

                return tableElements;
            };
         

        /// <summary>
        /// Do parsing
        /// </summary>
        public override XElement Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            this.ValidateStartTag(startElement, "Table");

            if (parentProcessor == null)
            {
                throw new ArgumentNullException();
            }

            var endTableTag = this.TryGetRequiredTags(startElement, "EndTable").First();
            var itemsSource = startElement.Value;            
            if (string.IsNullOrEmpty(itemsSource))
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, "Items"));
            }

            var tag = new TableTag
                          {
                              TagTable = startElement,
                              TagEndTable = endTableTag,
                              ItemsSource = itemsSource,
                          };
            tag.MakeTableElementCallback = MakeTableElementCallback;
            int? dynamicRow = null;
            var rowCount = 1;            
            var between = TraverseUtils.ElementsBetween(startElement, endTableTag).Descendants(WordMl.TableRowName);
            foreach (var tableRow in between)
            {
                if (tableRow.Descendants().Any(el => el.IsTag("ItemIndex") || el.IsTag("ItemText")))
                {
                    if (dynamicRow != null)
                    {
                        throw new Exception("Invalid template! Found several dynamic rows.");
                    }
                    dynamicRow = rowCount;                    
                }
                rowCount++;
            }
            var tableElement = startElement.NextSdt(WordMl.TableName).FirstOrDefault(element => element.IsBefore(endTableTag));
            if (tableElement != null)
            {
                tag.Table = tableElement;
            }

            tag.DynamicRow = dynamicRow;

            var processor = new TableProcessor { TableTag = tag };

            if (TraverseUtils.ElementsBetween(startElement, endTableTag).Any())
            {
                this.GoDeeper(processor, TraverseUtils.ElementsBetween(startElement, endTableTag).First());
            }
            parentProcessor.AddProcessor(processor);

            return endTableTag;
        }

        private static IEnumerable<TableElement> MakeTableElement(XElement startElement, XElement endElement)
        {
            ICollection<TableElement> tableElements = new Collection<TableElement>();
            var currentTagElement = TraverseUtils.NextTagElements(startElement).FirstOrDefault();
            while (currentTagElement != null && currentTagElement.IsBefore(endElement))
            {
                var tableElement = ToTableElement(currentTagElement);

                if (tableElement.IsItem || tableElement.IsIndex || tableElement.IsItemIf || tableElement.IsItemRepeater)
                {
                    tableElements.Add(tableElement);
                }
                if (tableElement.IsItemIf || tableElement.IsItemRepeater)
                {
                    currentTagElement = tableElement.EndTag;
                }
                currentTagElement = TraverseUtils.NextTagElements(currentTagElement).FirstOrDefault();
            }
            return tableElements;
        }

        private static TableElement ToTableElement(XElement tagElement)
        {
            var tableElement = new TableElement
                {
                    IsItem = tagElement.IsTag("itemtext"),
                    IsIndex = tagElement.IsTag("itemindex"),
                    IsItemIf = tagElement.IsTag("itemif"),
                    IsItemRepeater = tagElement.IsTag("itemrepeater"),
                    StartTag = tagElement,
                };
            if (tableElement.IsItem)
            {
                tableElement.Expression = tagElement.GetExpression();
            }
            else if (tableElement.IsItemIf || tableElement.IsItemRepeater)
            {
                tableElement.EndTag = FindEndTag(tableElement.StartTag, tableElement.IsItemRepeater ? "itemrepeater" : "itemif", tableElement.IsItemRepeater ? "enditemrepeater" : "enditemif");
                tableElement.Expression = tagElement.GetExpression();
                tableElement.TagElements = MakeTableElement(tableElement.StartTag, tableElement.EndTag);
            }

            return tableElement;
        }

        private static XElement FindEndTag(XElement startTag, string startTagName, string endTagName)
        {
            var startTagsOpened = 1;
            var current = startTag;
            while (startTagsOpened > 0 && current != null)
            {
                var nextTagElements = TraverseUtils.NextTagElements(current, new Collection<string> { startTagName, endTagName }).ToList();
                int index = -1;
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

        private void GoDeeper(ITagProcessor parentProcessor, XElement element)
        {
            do
            {
                if (element.IsSdt())
                {
                    switch (this.GetTagName(element).ToLower())
                    {
                        case "itemtext":
                        case "itemindex":
                        case "itemif":
                        case "itemrepeater":
                        case "enditemrepeater":
                        case "enditemif":
                            break;
                        default:
                            element = this.ParseSdt(parentProcessor, element);
                            break;
                    }
                }
                else if (element.HasElements)
                {
                    this.GoDeeper(parentProcessor, element.Elements().First());
                }
                element = element.NextElement();
            }
            while (element != null && (!element.IsSdt() || GetTagName(element).ToLower() != "endtable"));
        }
    }
}
