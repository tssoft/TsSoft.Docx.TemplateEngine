using System.Collections.Generic;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System;

    internal class RepeaterTag
    {
        public string Source { get; set; }

        public IEnumerable<RepeaterElement> Content { get; set; }

//        public XElement StartContent { get; set; }

//        public XElement EndContent { get; set; }

        public XElement StartRepeater { get; set; }

        public XElement EndRepeater { get; set; }

        public Func<XElement, RepeaterElement> MakeElementCallback { get; set; }
    }
}
