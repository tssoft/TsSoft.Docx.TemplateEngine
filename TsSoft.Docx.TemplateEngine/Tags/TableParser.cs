using System;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System.Collections.Generic;

    /// <summary>
    /// Parse table tag
    /// </summary>
    internal class TableParser : GeneralParser
    {
        /// <summary>
        /// Do parsing
        /// </summary>
        public override void Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            this.ValidateStartTag(startElement, "Table");

            if (parentProcessor == null)
            {
                throw new ArgumentNullException();
            }

            var endTableTag = this.TryGetRequiredTags(startElement, "EndTable").First();
            var itemsElements = this.TryGetRequiredTags(startElement, endTableTag, "Items");

            if (itemsElements.All(e => string.IsNullOrEmpty(e.Value)))
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, "Items"));
            }

            var contentElement = this.TryGetRequiredTags(startElement, endTableTag, "Content").First();
            var endContentElement = this.TryGetRequiredTags(contentElement, endTableTag, "EndContent").First();

            var tag = new TableTag
                          {
                              TagTable = startElement,
                              TagEndTable = endTableTag,
                              ItemsSource = itemsElements.First().Value,
                              TagContent = contentElement,
                              TagEndContent = endContentElement
                          };

            var dynamicRowElement = TraverseUtils.TagElementsBetween(startElement, endTableTag, "DynamicRow").FirstOrDefault();
            if (dynamicRowElement != null)
            {
                int dynamicRowValue;
                tag.DynamicRow = int.TryParse(dynamicRowElement.Value, out dynamicRowValue)
                                     ? dynamicRowValue
                                     : (int?)null;
            }

            var tableElement = contentElement.ElementsAfterSelf(WordMl.TableName).FirstOrDefault(element => element.IsBefore(endContentElement));
            if (tableElement != null)
            {
                tag.Table = tableElement;
            }

            var processor = new TableProcessor { TableTag = tag };
            parentProcessor.AddProcessor(processor);

            this.GoDeeper(processor, TraverseUtils.ElementsBetween(contentElement, endContentElement));
        }

        private void GoDeeper(ITagProcessor parentProcessor, IEnumerable<XElement> elements)
        {
            foreach (var element in elements)
            {
                if (element.IsSdt())
                {
                    switch (this.GetTagName(element).ToLower())
                    {
                        case "item":
                        case "itemindex":
                            continue;
                        default:
                            this.ParseSdt(parentProcessor, element);
                            break;
                    }
                }
                this.GoDeeper(parentProcessor, element.Elements());
            }
        }
    }
}
