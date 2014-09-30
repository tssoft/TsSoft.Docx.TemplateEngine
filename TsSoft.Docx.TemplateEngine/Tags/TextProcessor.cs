using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class TextProcessor<E> : ITagProcessor<E>
    {
        public void Process(TextTag textTag)
        {
            var element = textTag.TagNode;
            var text = DataReader.ReadText(textTag.Expression);
            element.AddBeforeSelf(DocxHelper.CreateTextElement(text));
            element.Remove();
        }
    }
}