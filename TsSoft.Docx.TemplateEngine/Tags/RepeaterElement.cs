using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class RepeaterElement
    {
        public IEnumerable<RepeaterElement> Elements { get; set; }

        public XElement XElement { get; set; }

        public IEnumerable<RepeaterElement> TagElements { get; set; }

        public string Expression { get; set; }

        public bool IsIndex { get; set; }

        public bool HasExpression
        {
            get
            {
                return !string.IsNullOrEmpty(this.Expression);
            }
        }

        public bool HasElements
        {
            get
            {
                return this.Elements != null && this.Elements.Any();
            }
        }

        public bool IsItem { get; set; }

        public bool IsItemIf { get; set; }

        public bool IsEndItemIf { get; set; }

        public bool IsItemRepeater { get; set; }

        public bool IsItemHtmlContent { get; set; }

        public bool IsItemTable { get; set; }

        public bool IsEndItemTable { get; set; }

        public XElement StartTag { get; set; }

        public XElement EndTag { get; set; }
    }
}
