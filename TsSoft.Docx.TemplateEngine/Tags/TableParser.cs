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
                throw new ArgumentNullException();
            }

            var endTableTag = TraverseUtils.NextTagElements(startElement, "EndTable").FirstOrDefault();
            if (endTableTag == null || TraverseUtils.TagElementsBetween(startElement, endTableTag, "Table").Any())
            {
                throw new Exception("Table closing tag wasn't found");
            }

            var tag = new TableTag();
            tag.TagTable = startElement;
            tag.TagEndTable = endTableTag;


            var itemsElement = TraverseUtils.TagElementsBetween(startElement, endTableTag, "Items").FirstOrDefault();
            if (itemsElement == null || itemsElement.Value == "")
            {
                throw new Exception("Table data source wasn't found");
            }
            tag.ItemsSource = itemsElement.Value;


            var dynamicRowElement = TraverseUtils.TagElementsBetween(startElement, endTableTag, "DynamicRow").FirstOrDefault();
            if (dynamicRowElement != null)
            {
                int dynamicRowValue;
                tag.DynamicRow = int.TryParse(dynamicRowElement.Value, out dynamicRowValue)
                                     ? dynamicRowValue
                                     : (int?) null;
            }


            var contentElement = TraverseUtils.TagElementsBetween(startElement, endTableTag, "Content").FirstOrDefault();
            if (contentElement == null)
            {
                throw new Exception("Context tag wasn't found");
            }

            tag.TagContent = contentElement;
            var endContentElement = TraverseUtils.TagElementsBetween(contentElement, endTableTag, "EndContent").FirstOrDefault();
            if (endContentElement == null)
            {
                throw new Exception("Context closing tag wasn't found");
            }
            tag.TagEndContent = endContentElement;

            var tableElement = contentElement.ElementsAfterSelf(WordMl.WordMlNamespace + "tbl").FirstOrDefault(element => element.IsBefore(endContentElement));
            if (tableElement != null)
            {
                tag.Table = tableElement;
            }

            var processor = new TableProcessor {TableTag = tag};
            parentProcessor.AddProcessor(processor);
        }
    }
}
