using System.Collections.Generic;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System.Linq;

    internal class TableElement
    {
        public bool IsItem { get; set; }

        public bool IsIndex { get; set; }

        public bool IsItemIf { get; set; }

        public bool IsItemHtmlContent { get; set; }

        public bool IsItemRepeater { get; set; }

        public bool IsItemTable { get; set; }

        public IEnumerable<TableElement> TagElements { get; set; }

        public string Expression { get; set; }

        public XElement StartTag { get; set; }

        public XElement EndTag { get; set; }

        public bool HasCell()
        {
            return this.StartTag.Descendants(WordMl.TableCellName).Any();
        }
    }
}
