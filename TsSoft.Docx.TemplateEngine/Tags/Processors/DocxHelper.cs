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

        private static readonly IEnumerable<XName> TextPropertiesNames = new Collection<XName>
                                                                                {

                                                                                };

        public static XElement CreateTextElement(XElement self, XElement parent, string text)
        {
            return CreateTextElement(self, parent, text, new XElement(WordMl.ParagraphName));
        }

        public static XElement CreateTextElement(XElement self, XElement parent, string text, XName wrapTo)
        {
            return CreateTextElement(self, parent, text, new XElement(wrapTo));
        }

        public static XElement CreateTextElement(XElement self, XElement parent, string text, XElement wrapTo)
        {
            var result = new XElement(WordMl.TextRunName, new XElement(WordMl.TextName, text));
            if (self.IsSdt())
            {
                var properties = self.Element(WordMl.SdtContentName)
                    .Descendants(WordMl.TextRunPropertiesName)
                    .Elements()
                    .Where(e => TextPropertiesNames.Contains(e.Name))
                    .Distinct(new NameComparer());
                var resultPropertiesTag = result.Elements(WordMl.TextRunPropertiesName).FirstOrDefault() ?? new XElement(WordMl.TextRunPropertiesName);
                resultPropertiesTag.Add(properties);
            }
            if (!ValidTextRunContainers.Any(name => name.Equals(parent.Name)))
            {
                wrapTo.Add(result);
                result = wrapTo;
            }
            return result;
        }

        private class NameComparer : IEqualityComparer<XElement>
        {
            public bool Equals(XElement x, XElement y)
            {
                return x.Name.Equals(y.Name);
            }

            public int GetHashCode(XElement obj)
            {
                return obj.Name.GetHashCode();
            }
        }
    }
}