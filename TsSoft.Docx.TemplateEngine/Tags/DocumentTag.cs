using System.Collections.Generic;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class DocumentTag
    {
        public IList<DocumentTag> InnerTags { get; private set; }

        public void Process()
        {
            foreach (var innerTag in InnerTags)
            {
                innerTag.Process();
            }
        }
    }
}