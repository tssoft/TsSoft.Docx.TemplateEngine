using System;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    public class LockValueAttribute : Attribute
    {
        public readonly string Value;

        public LockValueAttribute(string lockValue)
        {
            this.Value = lockValue;
        }
    }
}
