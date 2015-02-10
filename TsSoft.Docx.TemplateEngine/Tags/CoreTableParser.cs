using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class CoreTableParser
    {
        private string ItemIndexTagName;
        private string ItemTextTagName;
        private string ItemIfTagName;
        private string EndItemIfTagName;
        private string ItemHtmlContentTagName;

        private bool IsItemElement;

        public CoreTableParser(bool IsItemElement)
        {
            this.ItemIndexTagName = IsItemElement ? "titemindex" : "itemindex" ;
            this.ItemTextTagName = IsItemElement ? "titemtext" : "itemtext";
            this.ItemIfTagName = IsItemElement ? "titemif" : "itemif";
            this.EndItemIfTagName = IsItemElement ? "tenditemif" : "enditemif";
            this.ItemHtmlContentTagName = IsItemElement ? "titemhtmlcontent" : "itemhtmlcontent";
            this.IsItemElement = IsItemElement;
        }

        private static void ParseStructuredTableElement(TableElement tableElement, string startTagName, string endTagName)
        {
            tableElement.EndTag = FindEndTag(tableElement.StartTag, startTagName, endTagName);
            tableElement.Expression = tableElement.StartTag.GetExpression();
            tableElement.TagElements = MakeTableElement(tableElement.StartTag, tableElement.EndTag);
        }

        private static TableElement ToTableElement(XElement tagElement)
        {
            //TODO Make field constants
            var tableElement = new TableElement
            {
                IsItem = tagElement.IsTag("itemtext") || tagElement.IsTag("titemtext"),
                IsIndex = tagElement.IsTag("itemindex") || tagElement.IsTag("titemindex"),
                IsItemIf = tagElement.IsTag("itemif") || tagElement.IsTag("titemif"),
                IsItemRepeater = tagElement.IsTag("itemrepeater"),
                IsItemHtmlContent = tagElement.IsTag("itemhtmlcontent") || tagElement.IsTag("titemhtmlcontent"),
                StartTag = tagElement,
            };
            if (tableElement.IsItem || tableElement.IsItemHtmlContent)
            {
                tableElement.Expression = tagElement.GetExpression();
            }
           /* else if (tableElement.IsItemIf || tableElement.IsItemRepeater)
            {
                tableElement.EndTag = FindEndTag(tableElement.StartTag,
                                                 tableElement.IsItemRepeater ? "itemrepeater" : "itemif",
                                                 tableElement.IsItemRepeater ? "enditemrepeater" : "enditemif");
                tableElement.Expression = tagElement.GetExpression();
                tableElement.TagElements = MakeTableElement(tableElement.StartTag, tableElement.EndTag);
            }*/            
            else if (tableElement.IsItemHtmlContent)
            {
                tableElement.Expression = tagElement.GetExpression();
            }
            else if (tableElement.IsItemIf)
            {
                ParseStructuredTableElement(tableElement, "itemif", "enditemif");
            }
            else if (tableElement.IsItemRepeater)
            {
                ParseStructuredTableElement(tableElement, "itemrepeater", "enditemrepeater");
            }
            else if (tableElement.IsItemTable)
            {
                ParseStructuredTableElement(tableElement, "itemtable", "enditemtable");
            }
            return tableElement;
        }

        private Func<XElement, IEnumerable<TableElement>> MakeTableElementCallback = parentElement =>
        {
            ICollection<TableElement> tableElements = new Collection<TableElement>();
            var currentTagElement = TraverseUtils.NextTagElements(parentElement).FirstOrDefault();
            while (currentTagElement != null && currentTagElement.Ancestors().Any(element => element.Equals(parentElement)))
            {
                if (!currentTagElement.Name.Equals(WordMl.TableName))
                {
                    var tableElement = ToTableElement(currentTagElement);

                    if (tableElement.IsItem || tableElement.IsIndex || tableElement.IsItemIf ||
                        tableElement.IsItemRepeater || tableElement.IsItemHtmlContent || tableElement.IsItemTable)
                    {
                        tableElements.Add(tableElement);
                    }
                    if (tableElement.IsItemIf || tableElement.IsItemRepeater || tableElement.IsItemTable)
                    {
                        currentTagElement = tableElement.EndTag;
                    }
                    currentTagElement = TraverseUtils.NextTagElements(currentTagElement).FirstOrDefault();
                }
            }

            return tableElements;
        };

        public TableTag Parse(XElement startElement, XElement endTableTag)
        {
            var itemsSource = startElement.Value;
            if (String.IsNullOrEmpty(itemsSource))
            {
                throw new Exception(String.Format(MessageStrings.TagNotFoundOrEmpty, "Items"));
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
            // TODO Make several dynamic rows support
            // TODO Extend sdt names in loop
            foreach (var tableRow in between)
            {
                if (tableRow.Descendants().Any(el => el.IsTag(this.ItemIndexTagName) || el.IsTag(this.ItemTextTagName)))
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
            return tag;
        }        

        private static IEnumerable<TableElement> MakeTableElement(XElement startElement, XElement endElement)
        {
            ICollection<TableElement> tableElements = new Collection<TableElement>();
            var currentTagElement = TraverseUtils.NextTagElements(startElement).FirstOrDefault();
            while (currentTagElement != null && currentTagElement.IsBefore(endElement))
            {
                var tableElement = ToTableElement(currentTagElement);

                if (tableElement.IsItem || tableElement.IsIndex || tableElement.IsItemIf || tableElement.IsItemRepeater || tableElement.IsItemHtmlContent)
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
                throw new Exception(String.Format(MessageStrings.TagNotFoundOrEmpty, endTagName));
            }
            return current;
        }
    }
}
