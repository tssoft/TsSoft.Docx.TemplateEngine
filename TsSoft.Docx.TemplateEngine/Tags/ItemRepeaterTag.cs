using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class ItemRepeaterTag
    {
        public XElement StartItemRepeater { get; set; }

        public XElement EndItemRepeater { get; set; }

        public bool IsNotSeparatedRepeater { get; set; }

        public bool IsVisible { get; set; }

        public string Source { get; set; }

        public ICollection<ItemRepeaterTag> NestedRepeaters { get; set; }

        public Func<XElement, RepeaterElement> MakeElementCallback { get; set; }
    }
    
}
