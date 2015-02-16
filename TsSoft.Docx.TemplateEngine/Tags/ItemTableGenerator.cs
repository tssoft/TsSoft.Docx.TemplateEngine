using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class ItemTableGenerator
    {       
        public static XElement ProcessItemTableElement(XElement startTableElement, XElement endTableElement,
                                                       DataReader dataReader)
        {
            var tableElement = TraverseUtils.SecondElementsBetween(startTableElement,
                                                                          endTableElement)
                                                   .SingleOrDefault(re => re.Name.Equals(WordMl.TableName));
            var tableContainer = new XElement("TempContainerElement");
            tableContainer.Add(startTableElement);
            tableContainer.Add(tableElement);
            tableContainer.Add(endTableElement);
            var itemTableGenerator = new ItemTableGenerator();
            itemTableGenerator.Generate(tableContainer.Elements().First(), tableContainer.Elements().Last(), dataReader);
            return new XElement(tableContainer.Elements().SingleOrDefault());
        }
        public void Generate(XElement startTableElement, XElement endTableElement, DataReader dataReader)
        {
            var coreParser = new CoreTableParser(true);
            var tag = coreParser.Parse(startTableElement, endTableElement);
            var processor = new TableProcessor() { DataReader = dataReader, TableTag = tag };
            processor.Process();
        }
    }
}
