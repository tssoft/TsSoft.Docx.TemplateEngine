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
            var sdtElement = startElement.Element(WordMl.BodyName).Element(WordMl.SdtName);
            do
            {
                this.ParseSdt(parentProcessor, sdtElement);
                sdtElement = sdtElement.NextElement(x => x.Name == WordMl.SdtName);

            }
            while (sdtElement != null);
        }

        protected void ParseSdt(ITagProcessor parentProcessor, XElement sdtElement)
        {
            ITagParser parser = null;
            switch (GetTagName(sdtElement).ToLower())
            {
                case "text":
                    parser = new TextParser();
                    break;

                case "table":
                    parser = new TableParser();
                    break;

                case "repeater":
                    parser = new RepeaterParser();
                    break;

                case "if":
                    parser = new IfParser();
                    break;
            }
            if (parser != null)
            {
                parser.Parse(parentProcessor, sdtElement);
            }
        }

        protected void ValidateStartTag(XElement startElement, string tagName)
        {
            this.AssureElementExists(startElement, tagName);
            if (!startElement.Descendants(WordMl.TagName).Any(x => x.Attribute(WordMl.ValAttributeName).Value == tagName))
            {
                throw new Exception(MessageStrings.NotExpectedTag);
            }
        }

        protected XElement TryGetRequiredTag(XElement startElement, string tagName)
        {
            var tag = TraverseUtils.NextTagElements(startElement, tagName).SingleOrDefault();
            if (tag == null)
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, tagName));
            }
            return tag;
        }

        protected XElement TryGetRequiredTag(XElement startElement, XElement endElement, string tagName)
        {
            var tag = TraverseUtils.TagElementsBetween(startElement, endElement, tagName).SingleOrDefault();
            if (tag == null)
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, tagName));
            }
            return tag;
        }

        protected string GetTagName(XElement startElement)
        {
            var tagElement = startElement.Descendants(WordMl.TagName).SingleOrDefault();
            if (tagElement == null)
            {
                return null;
            }
            var nameAttribute = tagElement.Attribute(WordMl.ValAttributeName);
            return (nameAttribute == null) ? null : nameAttribute.Value;
        }

        private void AssureElementExists(XElement element, string tagName)
        {
            if (element == null)
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, tagName));
            }
        }
    }
}