using System;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Parsers
{
    /// <summary>
    /// Parse table tag
    /// </summary>
    /// <author>Георгий Поликарпов</author>
    class TableParser : GeneralParser
    {
        public Tags.TableTag Do(XElement startElement)
        {
            if (startElement == null)
            {
                throw new ArgumentNullException("Аргумент не может быть null");
            }

            var endTableTag = TraverseUtils.NextTagElements(startElement, "EndTable").FirstOrDefault();
            if (endTableTag == null || TraverseUtils.TagElementsBetween(startElement, endTableTag, "Table").Any())
            {
                throw new Exception("Отсутсвует закрывающий тег таблицы.");
            }

            var table = new Tags.TableTag();

            var itemsElement = TraverseUtils.TagElementsBetween(startElement, endTableTag, "Items").FirstOrDefault();
            if (itemsElement == null || itemsElement.Value == "")
            {
                throw new Exception("Отсутсвует источник данных для таблицы.");
            }
            table.ItemsSource = itemsElement.Value;

            var dynamicRowElement = TraverseUtils.TagElementsBetween(startElement, endTableTag, "DynamicRow").FirstOrDefault();
            if (dynamicRowElement != null)
            {
                table.DynamicRow = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            }

            var contentElement = TraverseUtils.TagElementsBetween(startElement, endTableTag, "Content").FirstOrDefault();
            if (contentElement == null)
            {
                throw new Exception("Отсутсвует тег контента.");
            }
            var endContentElement = TraverseUtils.TagElementsBetween(contentElement, endTableTag, "EndContent").FirstOrDefault();
            if (endContentElement == null || TraverseUtils.TagElementsBetween(contentElement, endContentElement, "Content").Any())
            {
                throw new Exception("Отсутсвует закрывающий тег контента.");
            }
            var tableElement = contentElement.ElementsAfterSelf(WordMl.WordMlNamespace + "tbl").Where(element => element.IsBefore(endContentElement)).FirstOrDefault();
            if (tableElement != null)
            {
                table.Table = tableElement;
            }

            return table;
        }

    }
}
