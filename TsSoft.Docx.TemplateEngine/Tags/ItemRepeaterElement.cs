using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class ItemRepeaterElement
    {
        public IEnumerable<ItemRepeaterElement> Elements { get; set; }

        public XElement XElement { get; set; }

        public string Expression { get; set; }

        public bool IsIndex { get; set; }

        public ItemRepeaterElement NextNestedRepeater { get; set; }

        public bool IsBeforeNestedRepeater
        {
            get { return this.NextNestedRepeater != null; }
        }

        public bool HasExpression
        {
            get { return !string.IsNullOrEmpty(this.Expression); }
        }

        public bool HasElements
        {
            get { return this.Elements != null && this.Elements.Any(); }
        }

        public XElement StartTag { get; set; }

        public XElement EndTag { get; set; }

        public bool IsItem { get; set; }

        public bool IsItemIf { get; set; }

        public bool IsVisible { get; set; }

        public bool IsEndItemIf { get; set; }

        public bool IsItemHtmlContent { get; set; }

        public bool IsItemRepeater { get; set; }

        public bool IsEndItemRepeater { get; set; }
    }   
}
