using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class GeneralParser : ITagParser
    {
        public virtual XElement Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            var sdtElement = startElement.Descendants(WordMl.SdtName).FirstOrDefault();
            while (sdtElement != null)
            {
                sdtElement = this.ParseSdt(parentProcessor, sdtElement);
                sdtElement = TraverseUtils.NextTagElements(sdtElement).FirstOrDefault();
            }
            return startElement;
        }

        protected XElement ParseSdt(ITagProcessor parentProcessor, XElement sdtElement)
        {
            ITagParser parser = null;
            switch (this.GetTagName(sdtElement).ToLower())
            {                
                case "htmlcontent":
                    parser = new HtmlContentParser();
                    break;

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
            return parser != null ? parser.Parse(parentProcessor, sdtElement) : sdtElement;
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
            var tag = TraverseUtils.NextTagElements(startElement, tagName).FirstOrDefault();
            if (tag == null)
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, tagName));
            }
            return tag;
        }

        protected XElement TryGetRequiredTag(XElement startElement, XElement endElement, string tagName)
        {
            var tag = TraverseUtils.TagElementsBetween(startElement, endElement, tagName).FirstOrDefault();
            if (tag == null)
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, tagName));
            }
            return tag;
        } 
        
        protected ICollection<XElement> TryGetRequiredTags(XElement startElement, XElement endElement, string tagName)
        {
            var tags = TraverseUtils.TagElementsBetween(startElement, endElement, tagName).ToList();
            if (!tags.Any())
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, tagName));
            }
            return tags;
        } 
    
        protected ICollection<XElement> TryGetRequiredTags(XElement startElement, string tagName)
        {
            var tags = TraverseUtils.NextTagElements(startElement, tagName).ToList();
            if (!tags.Any())
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, tagName));
            }
            return tags;
        }

        protected string GetTagName(XElement startElement)
        {
            if (startElement == null)
            {
                throw new ArgumentNullException(string.Format("Argument {0} cannot be null", "startElement"));
            }
            var tagElement = startElement.Descendants(WordMl.TagName).SingleOrDefault();
            if (tagElement == null)
            {
                return null;
            }
            var nameAttribute = tagElement.Attribute(WordMl.ValAttributeName);
            return (nameAttribute == null) ? null : nameAttribute.Value;
        }

// ReSharper disable UnusedParameter.Local
        private void AssureElementExists(XElement element, string tagName)
// ReSharper restore UnusedParameter.Local
        {
            if (element == null)
            {
                throw new Exception(string.Format(MessageStrings.TagNotFoundOrEmpty, tagName));
            }
        }
    }
}