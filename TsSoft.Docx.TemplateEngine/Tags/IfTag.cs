using System.Collections.Generic;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System.Xml.Linq;

    internal class IfTag
    {
        public XElement StartIf { get; set; }

        public XElement EndIf { get; set; }

        public IEnumerable<XElement> IfContent { get; set; }
    }
}
