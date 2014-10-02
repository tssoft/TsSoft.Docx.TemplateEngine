using System;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class GeneralParser
    {
        public virtual ITagProcessor Parse(XElement startElement)
        {
            return null;
        }

        protected void ValidateStartTag(XElement startElement, string tagName)
        {
            if (startElement == null)
            {
                // TODO
                throw new Exception();
            }
            if (startElement.Name != WordMl.SdtName)
            {
                // TODO
                throw new Exception();
            }
            if (null == startElement.Descendants(WordMl.TagName).SingleOrDefault(x => x.Attribute(WordMl.ValAttributeName).Value == tagName))
            {
                // TODO
                throw new Exception();
            }
        }
    }
}