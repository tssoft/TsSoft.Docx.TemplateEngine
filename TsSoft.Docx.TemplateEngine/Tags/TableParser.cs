using System;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    /// <summary>
    /// Parse table tag
    /// </summary>
    class TableParser : GeneralParser
    {
        /// <summary>
        /// Do parsing
        /// </summary>
        public override void Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            ValidateStartTag(startElement, "Table");

            if (parentProcessor == null)
            {
                throw new ArgumentNullException(string.Format(MessageStrings.ArgumentNull, "parentProcessor"));
            }

            var endTableElement = TagElement(startElement, "EndTable");
            if (endTableElement == null || TagElementBetween(startElement, endTableElement, "Table") != null)
            {
                throw new Exception("Table closing tag wasn't found");
            }

            var tag = new TableTag();
            tag.TagTable = startElement;
            tag.TagEndTable = endTableElement;

            var itemsElement = TagElementBetween(startElement, endTableElement, "Items");
            if (itemsElement == null || itemsElement.Value == string.Empty)
            {
                throw new Exception("Table data source wasn't found");
            }
            tag.ItemsSource = itemsElement.Value;

            var dynamicRowElement = TagElementBetween(startElement, endTableElement, "DynamicRow");
            if (dynamicRowElement != null)
            {
                int dynamicRowValue;
                tag.DynamicRow = int.TryParse(dynamicRowElement.Value, out dynamicRowValue)
                                     ? dynamicRowValue
                                     : (int?) null;
            }

            var contentElement = TagElementBetween(startElement, endTableElement, "Content");
            if (contentElement == null)
            {
                throw new Exception("Context tag wasn't found");
            }
            tag.TagContent = contentElement;
            var endContentElement = TagElementBetween(contentElement, endTableElement, "EndContent");
            if (endContentElement == null)
            {
                throw new Exception("Context closing tag wasn't found");
            }
            tag.TagEndContent = endContentElement;

            var tableElement = contentElement.ElementsAfterSelf(WordMl.TableName).FirstOrDefault(element => element.IsBefore(endContentElement));
            if (tableElement != null)
            {
                tag.Table = tableElement;
            }

            var processor = new TableProcessor {TableTag = tag};
            parentProcessor.AddProcessor(processor);
        }

        private XElement TagElement(XElement startElement, string tagName)
        {
            return
                startElement.ElementsAfterSelf(WordMl.SdtName)
                            .FirstOrDefault(
                                element =>
                                element.Element(WordMl.SdtPrName)
                                       .Element(WordMl.TagName)
                                       .Attribute(WordMl.ValAttributeName)
                                       .Value == tagName);
        }

        private XElement TagElementBetween(XElement startElement, XElement endElement, string tagName)
        {
            return
                startElement.ElementsAfterSelf(WordMl.SdtName)
                            .FirstOrDefault(
                                element =>
                                element.Element(WordMl.SdtPrName)
                                       .Element(WordMl.TagName)
                                       .Attribute(WordMl.ValAttributeName)
                                       .Value == tagName
                                && element.IsBefore(endElement));
        }
    }
}
