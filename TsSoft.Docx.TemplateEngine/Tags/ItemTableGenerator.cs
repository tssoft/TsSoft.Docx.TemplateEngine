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
        public void Generate(XElement startTableElement, XElement endTableElement, DataReader dataReader)
        {
            var coreParser = new CoreTableParser(true);
            var tag = coreParser.Parse(startTableElement, endTableElement);
            var processor = new TableProcessor() { DataReader = dataReader, TableTag = tag };
            processor.Process();
        }
    }
}
