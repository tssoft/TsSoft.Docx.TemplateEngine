using System;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine
{
    internal static class XElementExtension
    {
        public static XElement NextElement(this XElement element)
        {
            return element.NodesAfterSelf().OfType<XElement>().FirstOrDefault();
        }

        public static XElement NextElement(this XElement element, Func<XElement, bool> predicate)
        {
            return element.NodesAfterSelf().OfType<XElement>().FirstOrDefault(predicate);
        }
    }
}