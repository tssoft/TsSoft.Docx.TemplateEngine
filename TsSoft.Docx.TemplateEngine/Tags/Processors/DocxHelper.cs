using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;

    internal class DocxHelper
    {
        private static readonly IEnumerable<XName> ValidTextRunContainers = new Collection<XName>
                                                                                {
                                                                                    WordMl.ParagraphName,
                                                                                    WordMl.CustomXmlName,
                                                                                    WordMl.HyperlinkName,
                                                                                    WordMl.MoveFromName,
                                                                                    WordMl.MoveToName,
                                                                                    WordMl.SdtContentName,
                                                                                    WordMl.DeleteName,
                                                                                    WordMl.InsertName,
                                                                                    WordMl.SmartTagName,
                                                                                    WordMl.FieldSimpleName
                                                                                };

        public static XElement CreateTextElement(XElement parent, string text)
        {
            return CreateTextElement(parent, text, new XElement(WordMl.ParagraphName));
        }

        public static XElement CreateTextElement(XElement parent, string text, XName wrapTo)
        {
            return CreateTextElement(parent, text, new XElement(wrapTo));
        }

        public static XElement CreateTextElement(XElement parent, string text, XElement wrapTo)
        {
            var result = new XElement(WordMl.TextRunName, new XElement(WordMl.TextName, text));
            if (!ValidTextRunContainers.Any(name => name.Equals(parent.Name)))
            {
                wrapTo.Add(result);
                result = wrapTo;
            }
            return result;
        }
    }
}