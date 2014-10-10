using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{

    internal class TableElement
    {
        public bool IsItem { get; set; }

        public bool IsIndex { get; set; }

        public bool IsItemIf { get; set; }

        public XElement XElement { get; set; }

        public IEnumerable<TableElement> TagElements { get; set; }

        public string Expression { get; set; }

        public XElement StartTag { get; set; }

        public XElement EndTag { get; set; }
    }
}
