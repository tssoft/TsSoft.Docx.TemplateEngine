using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Parsers
{
    internal class GeneralParser
    {
        protected void ValidateStartTag(XElement startElement, string tagName)
        {
            AssureElementExists(startElement);
            if (null == startElement.Descendants(WordMl.TagName).SingleOrDefault(x => x.Attribute(WordMl.ValAttributeName).Value == tagName))
            {
                // TODO
                throw new Exception();
            }
        }

        protected XElement TryGetRequiredTag(XElement startElement, string tagName)
        {
            var tag = TraverseUtils.NextTagElements(startElement, tagName).SingleOrDefault();
            if (tag == null)
            {
                throw new Exception(String.Format("Required tag {0} not found", tagName));
            }
            return tag;
        }

        protected XElement TryGetRequiredTag(XElement startElement, XElement endElement, string tagName)
        {
            var tag = TraverseUtils.TagElementsBetween(startElement, endElement, tagName).SingleOrDefault();
            if (tag == null)
            {
                throw new Exception(String.Format("Required tag {0} not found", tagName));
            }
            return tag;
        }

        private void AssureElementExists(XElement element)
        {
            if (element == null)
            {
                // TODO
                throw new Exception();
            }
            if (element.Name != WordMl.SdtName)
            {
                // TODO
                throw new Exception();
            }
        }
    }
}