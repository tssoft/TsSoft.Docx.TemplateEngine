using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class RepeaterElement
    {
        public IEnumerable<RepeaterElement> Elements { get; set; }

        public XElement XElement { get; set; }

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
    }
}
