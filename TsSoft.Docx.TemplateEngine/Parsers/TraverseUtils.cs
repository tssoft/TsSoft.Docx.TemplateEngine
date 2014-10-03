using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Parsers
{
    internal static class TraverseUtils
    {

        public static IEnumerable<XElement> NextTagElements(XElement startElement, string tagName)
        {
            return startElement.ElementsAfterSelf(WordMl.SdtName)
                .Where(element => element.Element(WordMl.SdtPrName)
                .Element(WordMl.TagName).Attribute(WordMl.ValAttributeName).Value.Equals(tagName));
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
            return root.Elements(WordMl.SdtName).Where(element => element.Element(WordMl.SdtPrName).Element(WordMl.TagName).Attribute(WordMl.ValAttributeName).Value == tagName).First();
        } 

        public static bool IsTag(this XElement self, String tagName)
        {
            return self.Descendants(WordMl.SdtPrName)
                .Descendants(WordMl.TagName)
                .SingleOrDefault(e => e.Attribute(WordMl.ValAttributeName).Value.Equals(tagName)) != null;
        } 
        
        public static bool IsSdt(this XElement self)
        {
            return self.Name.Equals(WordMl.SdtName);
        }

        public static String GetExpression(this XElement self)
        {
            return self.IsSdt() ? self.Value : null;
        }
    }
}
