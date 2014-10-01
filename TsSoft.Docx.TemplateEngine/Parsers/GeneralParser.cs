using System;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Parsers
{
    internal class GeneralParser
    {
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