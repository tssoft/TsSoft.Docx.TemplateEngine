using System;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class GeneralParser : ITagParser
    {
        public virtual void Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            XElement sdtElement = startElement.NextElement(x => x.Name == WordMl.SdtName);
            while (sdtElement != null)
            {
                ParseSdt(parentProcessor, sdtElement);
            }
        }

        protected void ParseSdt(ITagProcessor parentProcessor, XElement sdtElement)
        {
            ITagParser parser = null;
            // TODO Ignore case
            switch (GetTagName(sdtElement))
            {
                case "Text":
                    parser = new TextParser();
                    break;

                case "Table":
                    break;

                case "Repeater":
                    break;

                case "If":
                    break;
            }
            if (parser != null)
            {
                parser.Parse(parentProcessor, sdtElement);
            }
        }

        protected void ValidateStartTag(XElement startElement, string tagName)
        {
            AssureElementExists(startElement);
            if (null == startElement.Descendants(WordMl.TagName).SingleOrDefault(x => x.Attribute(WordMl.ValAttributeName).Value == tagName))
            {
                // TODO
                throw new Exception(MessageStrings.NotExpectedTag);
            }
        }

        protected XElement TryGetRequiredTag(XElement startElement, string tagName)
        {
            var tag = TraverseUtils.NextTagElements(startElement, tagName).SingleOrDefault();
            if (tag == null)
            {
                throw new Exception(String.Format(MessageStrings.TagNotFoundOrEmpty, tagName));
            }
            return tag;
        }

        protected XElement TryGetRequiredTag(XElement startElement, XElement endElement, string tagName)
        {
            var tag = TraverseUtils.TagElementsBetween(startElement, endElement, tagName).SingleOrDefault();
            if (tag == null)
            {
                throw new Exception(String.Format(MessageStrings.TagNotFoundOrEmpty, tagName));
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

        protected string GetTagName(XElement startElement)
        {
            var tagElement = startElement.Descendants(WordMl.TagName).SingleOrDefault();
            if (tagElement == null)
            {
                return null;
            }
            var nameAttribute = tagElement.Attribute(WordMl.ValAttributeName);
            return (nameAttribute != null) ? nameAttribute.Value : null;
        }
    }
}