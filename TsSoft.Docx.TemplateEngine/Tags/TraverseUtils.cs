using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System;
    using System.Globalization;

    internal static class TraverseUtils
    {
        public static IEnumerable<XElement> NextTagElements(XElement startElement)
        {
            return NextTagElements(startElement, Enumerable.Empty<string>());
        }

        public static IEnumerable<XElement> NextTagElements(XElement startElement, string tagName)
        {
            return NextTagElements(startElement, new List<string> { tagName });
        }

        public static IEnumerable<XElement> NextTagElements(XElement startElement, IEnumerable<string> tagNames)
        {
            var namesList = tagNames.Select(s => s.ToLower()).ToList();
            var nextTagElements = startElement.ElementsAfterSelf(WordMl.SdtName).Where(MakeTagMatchingPredicate(namesList));
            if (nextTagElements.FirstOrDefault() == null)
            {
                var nextTagsInnerElements = GoDeeper(startElement, namesList);
                return nextTagsInnerElements.FirstOrDefault() != null ? nextTagsInnerElements : GoUp(startElement, namesList);
            }
            return nextTagElements;
        }

        public static IEnumerable<XElement> TagElementsBetween(XElement startElement, XElement endElement, string tagName)
        {
            return NextTagElements(startElement, tagName).Where(element => element.IsBefore(endElement));
        }

        public static IEnumerable<XElement> ElementsBetween(XElement startElement, XElement endElement)
        {
            if (startElement.Parent.Equals(endElement.Parent))
            {
                return startElement.ElementsAfterSelf().Where(e => e.IsBefore(endElement));
            }
            var commonParent = startElement.Ancestors().Intersect(endElement.Ancestors()).First();
            var endElementFirstLevel = commonParent.Elements().First(e => e.Descendants().Contains(endElement));
            var startElementFirstLevel = commonParent.Elements().First(e => e.Descendants().Contains(startElement));
            var afterStart = startElement.ElementsAfterSelf();
            var between = ElementsBetween(startElementFirstLevel, endElementFirstLevel);
            var beforeEnd = endElementFirstLevel.DescendantsBefore(endElement);
            return afterStart.Union(between).Union(beforeEnd);
        }

        public static XElement TagElement(XElement root, string tagName)
        {
            return root.Descendants(WordMl.SdtName).First(element => element.Element(WordMl.SdtPrName).Element(WordMl.TagName).Attribute(WordMl.ValAttributeName).Value == tagName);
        }

        public static bool IsTag(this XElement self, string tagName)
        {
            return self.IsSdt()
                ? self.Elements(WordMl.SdtPrName).Elements(WordMl.TagName).SingleOrDefault(e => e.Attribute(WordMl.ValAttributeName).Value.ToLower(CultureInfo.CurrentCulture)
                    .Equals(tagName.ToLower(CultureInfo.CurrentCulture))) != null
                : false;
        }

        public static bool IsSdt(this XElement self)
        {
            return self.Name.Equals(WordMl.SdtName);
        }

        /// <summary>
        /// If you imagine an XML tree so that the root is the upmost element, than you can say
        /// that this method recursively takes all descendant elements of "self" to the right of "element"
        /// </summary>
        /// <param name="self"></param>
        /// <param name="element"></param>
        /// <returns></returns>
        public static IEnumerable<XElement> DescendantsBefore(this XElement self, XElement element)
        {
            if (self.Elements().Contains(element))
            {
                return element.ElementsBeforeSelf();
            }
            var container = self.Elements().First(e => e.Descendants().Contains(element));
            return self.Elements().Where(e => e.IsBefore(container)).Union(container.DescendantsBefore(element));
        }

        public static string GetExpression(this XElement self)
        {
            return self.IsSdt() ? self.Value : null;
        }

        private static Func<XElement, bool> MakeTagMatchingPredicate(ICollection<string> tagNames)
        {
            Func<XElement, bool> func;
            if (!tagNames.Any())
            {
                func = element => true;
            }
            else
            {
                func = element => tagNames.Contains(element.Element(WordMl.SdtPrName).Element(WordMl.TagName).Attribute(WordMl.ValAttributeName).Value.ToLower());
            }
            return func;
        }

        private static IEnumerable<XElement> GoDeeper(XElement startElement, ICollection<string> tagNames)
        {
            return startElement.ElementsAfterSelf().Descendants(WordMl.SdtName).Where(MakeTagMatchingPredicate(tagNames));
        }

        private static IEnumerable<XElement> GoUp(XElement startElement, ICollection<string> tagNames)
        {
            IEnumerable<XElement> result = Enumerable.Empty<XElement>();
            var parent = startElement.Parent;
            while (result.FirstOrDefault() == null && !parent.Name.Equals(WordMl.BodyName))
            {
                result = parent.ElementsAfterSelf().DescendantsAndSelf(WordMl.SdtName).Where(MakeTagMatchingPredicate(tagNames));
                parent = parent.Parent;
            }
            return result;
        }
    }
}
