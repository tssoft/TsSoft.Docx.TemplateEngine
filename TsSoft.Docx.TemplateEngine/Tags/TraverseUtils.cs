using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System;
    using System.Globalization;

    internal static class TraverseUtils
    {

        public static IEnumerable<XElement> NextTags(this XElement self)
        {
            return NextTagElements(self);
        }

        public static IEnumerable<XElement> NextTags(this XElement self, string tagName)
        {
            return NextTagElements(self, tagName);
        }

        public static IEnumerable<XElement> NextSdt(this XElement self, IEnumerable<XName> names)
        {
            return NextSdtElements(self, names);
        }

        public static IEnumerable<XElement> NextSdt(this XElement self)
        {
            return NextSdtElements(self, Enumerable.Empty<XName>());
        }

        public static IEnumerable<XElement> NextSdt(this XElement self, XName name)
        {
            return NextSdtElements(self, new List<XName> { name });
        }

        public static IEnumerable<XElement> NextTags(this XElement self, IEnumerable<string> tagNames)
        {
            return NextTagElements(self, tagNames);
        }

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
            var nextTagsInnerElements = GoDeeper(startElement, namesList);
            if (nextTagsInnerElements.FirstOrDefault() == null)
            {
                var nextTagElements = startElement.ElementsAfterSelf(WordMl.SdtName).Where(MakeTagMatchingPredicate(namesList));
                return nextTagElements.FirstOrDefault() != null ? nextTagElements : GoUp(startElement, namesList);
            }
            return nextTagsInnerElements;
        }

        public static IEnumerable<XElement> NextSdtElements(XElement startElement, IEnumerable<XName> names)
        {
            var namesList = names.ToList();
            var nextTagsInnerElements = GoDeeper(startElement, namesList);
            if (nextTagsInnerElements.FirstOrDefault() == null)
            {
                var nextTagElements = startElement.ElementsAfterSelf().Where(MakeSdtMatchingPredicate(namesList));
                return nextTagElements.FirstOrDefault() != null ? nextTagElements : GoUp(startElement, namesList);
            }
            return nextTagsInnerElements;
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
            var commonParent = startElement.Ancestors().Intersect(endElement.Ancestors()).FirstOrDefault();
            if (commonParent == null)
            {
                return Enumerable.Empty<XElement>();
            }
            var endElementFirstLevel = commonParent.Elements().First(e => e.Descendants().Contains(endElement) || e.Equals(endElement));
            var startElementFirstLevel = commonParent.Elements().First(e => e.Descendants().Contains(startElement) || e.Equals(startElement));
            var afterStart = startElement.ElementsAfterSelf();
            var between = ElementsBetween(startElementFirstLevel, endElementFirstLevel);
            var beforeEnd = endElementFirstLevel.DescendantsBefore(endElement);
            return afterStart.Union(between).Union(beforeEnd);
        }
        // TODO make one second elements between       
        public static IEnumerable<XElement> SecondElementsBetween(XElement startElement, XElement endElement, bool nested = false)
        {
            if (startElement.Parent.Equals(endElement.Parent))
            {
                var oneLevelElements = new List<XElement>();
                if (nested)
                {
                    oneLevelElements.Add(startElement);
                }
                oneLevelElements.AddRange(startElement.ElementsAfterSelf().Where(e => e.IsBefore(endElement)));                
                return oneLevelElements.AsEnumerable();               
            }
            XElement currentElement = startElement.NextElementWithUpTransition();
            var elements = new List<XElement>();
            while ((currentElement != null) && (currentElement != endElement))
            {
                if (currentElement.HasElements && currentElement.Descendants().Contains(endElement))
                {
                    elements.AddRange(SecondElementsBetween(currentElement.Elements().First(), endElement, true));
                    break;                    
                }
                elements.Add(currentElement);
                currentElement = currentElement.NextElementWithUpTransition();
            }
            return elements;
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

        public static string GenerateRandomRsidR()
        {
            const int SixByteDecimalMax = 0x00FFFFFF;
            const int SixByteDecimalMin = 0x00111111;
            var resultString = "00";
            var random = new Random();
            resultString += random.Next(SixByteDecimalMin, SixByteDecimalMax).ToString("X");
            return resultString;
        } 
        
        public static XElement NextElementWithUpTransition(this XElement self)
        {
            var nextElement = self.NextElement();
            return (nextElement == null && (self.Parent != null) && (self.Parent.NextElement() != null))
                       ? self.Parent.NextElement()
                       : nextElement;
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
            if (self.Equals(element))
            {
                return Enumerable.Empty<XElement>();
            }
            if (self.Elements().Contains(element))
            {
                return element.ElementsBeforeSelf();
            }
            // --
            var container = self.Elements().First(e => e.Descendants().Contains(element) || e.Equals(element));

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
        private static Func<XElement, bool> MakeSdtMatchingPredicate(ICollection<XName> names)
        {
            Func<XElement, bool> func;
            if (!names.Any())
            {
                func = element => true;
            }
            else
            {
                func = element => names.Contains(element.Name);
            }
            return func;
        }

        private static IEnumerable<XElement> GoDeeper(XElement startElement, ICollection<string> tagNames)
        {
            var tagsElements = startElement.Descendants(WordMl.SdtName).Where(MakeTagMatchingPredicate(tagNames));
            bool IsAnyTagsElement = tagsElements.Any();

            return IsAnyTagsElement
                       ? tagsElements
                       : startElement.ElementsAfterSelf()
                                     .DescendantsAndSelf(WordMl.SdtName)
                                     .Where(MakeTagMatchingPredicate(tagNames));

        }

        private static IEnumerable<XElement> GoDeeper(XElement startElement, ICollection<XName> names)
        {
            var tagsElements = startElement.Descendants().Where(MakeSdtMatchingPredicate(names));
            bool IsAnyTagsElement = tagsElements.Any();

            return IsAnyTagsElement
                       ? tagsElements
                       : startElement.ElementsAfterSelf()
                                     .DescendantsAndSelf()
                                     .Where(MakeSdtMatchingPredicate(names));

        }

        private static IEnumerable<XElement> GoUp(XElement startElement, ICollection<string> tagNames)
        {
            IEnumerable<XElement> result = Enumerable.Empty<XElement>();
            var parent = startElement.Parent;
            while (parent != null && result.FirstOrDefault() == null && !parent.Name.Equals(WordMl.BodyName))
            {
                result = parent.ElementsAfterSelf().DescendantsAndSelf(WordMl.SdtName).Where(MakeTagMatchingPredicate(tagNames));
                parent = parent.Parent;
            }
            return result;
        }

        private static IEnumerable<XElement> GoUp(XElement startElement, ICollection<XName> names)
        {
            IEnumerable<XElement> result = Enumerable.Empty<XElement>();
            var parent = startElement.Parent;
            while (parent != null && result.FirstOrDefault() == null && !parent.Name.Equals(WordMl.BodyName))
            {
                result = parent.ElementsAfterSelf().DescendantsAndSelf().Where(MakeSdtMatchingPredicate(names));
                parent = parent.Parent;
            }
            return result;
        }               

    }
}
