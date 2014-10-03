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
            if (startElement == null)
            {
                // TODO
                throw new ArgumentNullException(string.Format(MessageStrings.ArgumentNull, "startElement"));
            }
            if (startElement.Name != WordMl.SdtName)
            {
                // TODO
                throw new Exception(MessageStrings.NotATag);
            }
            if (null == startElement.Descendants(WordMl.TagName).SingleOrDefault(x => x.Attribute(WordMl.ValAttributeName).Value == tagName))
            {
                // TODO
                throw new Exception(MessageStrings.NotExpectedTag);
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