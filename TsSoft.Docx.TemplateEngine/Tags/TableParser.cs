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
        public override XElement Parse(ITagProcessor parentProcessor, XElement startElement)
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

            var contentTag = this.TryGetRequiredTags(startElement, endTableTag, "Content").First();
            var endContentTag = this.TryGetRequiredTags(contentTag, endTableTag, "EndContent").First();

            var tag = new TableTag
                          {
                              TagTable = startElement,
                              TagEndTable = endTableTag,
                              ItemsSource = itemsElements.First().Value,
                              TagContent = contentTag,
                              TagEndContent = endContentTag
                          };

            var dynamicRowElement = TraverseUtils.TagElementsBetween(startElement, endTableTag, "DynamicRow").FirstOrDefault();
            if (dynamicRowElement != null)
            {
                int dynamicRowValue;
                tag.DynamicRow = int.TryParse(dynamicRowElement.Value, out dynamicRowValue)
                                     ? dynamicRowValue
                                     : (int?)null;
            }

            var tableElement = contentTag.ElementsAfterSelf(WordMl.TableName).FirstOrDefault(element => element.IsBefore(endContentTag));
            if (tableElement != null)
            {
                tag.Table = tableElement;
            }

            var processor = new TableProcessor { TableTag = tag };

            var contentElements = TraverseUtils.ElementsBetween(contentTag, endContentTag);
            if (contentElements.Any())
            {
                this.GoDeeper(processor, contentElements.First());
            }
            parentProcessor.AddProcessor(processor);

            return endTableTag;
        }

        private void GoDeeper(ITagProcessor parentProcessor, XElement element)
        {
            do
            {
                if (element.IsSdt())
                {
                    switch (this.GetTagName(element).ToLower())
                    {
                        case "item":
                        case "itemindex":
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
            while (element != null && (!element.IsSdt() || GetTagName(element).ToLower() != "endcontent"));
        }
    }
}
