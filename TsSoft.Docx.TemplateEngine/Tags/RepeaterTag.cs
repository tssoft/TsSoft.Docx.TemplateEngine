
using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Xml.XPath;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class RepeaterTag
    {
        public String Source { get; set; }

        public IEnumerable<RepeaterElement> Content { get; set; } 
        public XElement Start { get; set; }
        public XElement End { get; set; }
    }
}
