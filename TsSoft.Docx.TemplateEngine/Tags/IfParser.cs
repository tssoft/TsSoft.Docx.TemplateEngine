using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System.Xml.Linq;

    using TsSoft.Docx.TemplateEngine.Tags.Processors;

    internal class IfParser : GeneralParser
    {
        public override void Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            base.Parse(parentProcessor, startElement);
        }
    }
}
