using System;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System.Xml.Linq;

    public class DynamicContentLocker
    {
        private readonly XName lokingElementName = WordMl.WordMlNamespace + "lock";

        public XElement GetLockingElement(SdtTagLockingType lockingType)
        {
            return new XElement(
                this.lokingElementName,
                new XAttribute(WordMl.ValAttributeName, this.GetLockingValue(lockingType)));
        }

        private string GetLockingValue(SdtTagLockingType lockingType)
        {
            var attributeType = lockingType.GetType().GetField(lockingType.ToString());
            var attribute = (LockValueAttribute)Attribute.GetCustomAttribute(attributeType, typeof(LockValueAttribute));
            return attribute.Value;
        }
    }
}
