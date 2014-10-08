using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System;

    internal static class TraverseUtils
    {
        public static IEnumerable<XElement> NextTagElements(XElement startElement)
        {
            return NextTagElements(startElement, null);
        }

        public static IEnumerable<XElement> NextTagElements(XElement startElement, string tagName)
        {
            var nextTagElements = startElement.ElementsAfterSelf(WordMl.SdtName).Where(MakeTagMatchingPredicate(tagName)).ToList();
            if (!nextTagElements.Any())
            {
                var nextTagsInnerElements = GoDeeper(startElement, tagName).ToList();
                return nextTagsInnerElements.Any() ? nextTagsInnerElements : GoUp(startElement, tagName);
            }
            return nextTagElements;
        }

        public static IEnumerable<XElement> TagElementsBetween(XElement startElement, XElement endElement, string tagName)
        {
            return NextTagElements(startElement, tagName).Where(element => element.IsBefore(endElement));
        }

        public static IEnumerable<XElement> ElementsBetween(XElement startElement, XElement endElement)
        {
            return startElement.ElementsAfterSelf().Where(e => e.IsBefore(endElement));
        }

        public static XElement TagElement(XElement root, string tagName)
        {
            return root.Descendants(WordMl.SdtName).First(element => element.Element(WordMl.SdtPrName).Element(WordMl.TagName).Attribute(WordMl.ValAttributeName).Value == tagName);
        }

        public static bool IsTag(this XElement self, string tagName)
        {
            return self.IsSdt() ? self.Elements(WordMl.SdtPrName)
                .Elements(WordMl.TagName)
                .SingleOrDefault(e => e.Attribute(WordMl.ValAttributeName).Value.Equals(tagName)) != null : false;
        }

        public static bool IsSdt(this XElement self)
        {
            return self.Name.Equals(WordMl.SdtName);
        }

        public static string GetExpression(this XElement self)
        {
            return self.IsSdt() ? self.Value : null;
        }

        private static Func<XElement, bool> MakeTagMatchingPredicate(string tagName)
        {
            Func<XElement, bool> func;
            if (string.IsNullOrEmpty(tagName))
            {
                func = element => true;
            }
            else
            {
                func = element => element.Element(WordMl.SdtPrName).Element(WordMl.TagName).Attribute(WordMl.ValAttributeName).Value.Equals(tagName);
            }
            return func;
        }

        private static IEnumerable<XElement> GoDeeper(XElement startElement, string name)
        {
            return startElement.ElementsAfterSelf().Descendants(WordMl.SdtName).Where(MakeTagMatchingPredicate(name));
        }

        private static IEnumerable<XElement> GoUp(XElement startElement, string name)
        {
            IEnumerable<XElement> result = Enumerable.Empty<XElement>();
            var parent = startElement.Parent;
            while (!result.Any() && !parent.Name.Equals(WordMl.BodyName))
            {
                result = parent.ElementsAfterSelf().DescendantsAndSelf(WordMl.SdtName).Where(MakeTagMatchingPredicate(name));
                parent = parent.Parent;
            }
            return result;
        }
    }
}
