
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class RepeaterElement
    {
        public IEnumerable<RepeaterElement> Elements { get; set; }

        public XElement XElement { get; set; }

        public String Expression { get; set; }

        public Boolean IsIndex { get; set; }

        public Boolean HasExpression { get { return !String.IsNullOrEmpty(Expression); } }

        public Boolean HasElements { get { return Elements != null && Elements.Any(); } }
        public Boolean IsItem { get; set; }

    }
}
