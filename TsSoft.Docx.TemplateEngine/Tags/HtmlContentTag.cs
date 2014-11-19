using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class HtmlContentTag
    {
        public string Expression { get; set; }

        public XElement TagNode { get; set; }
    }
}
