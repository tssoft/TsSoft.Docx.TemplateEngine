namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    using System.Linq;
    using System.Xml.Linq;

    internal class TextProcessor : AbstractProcessor
    {
        public TextTag TextTag { get; set; }

        public override void Process()
        {
            base.Process();
            var element = TextTag.TagNode;
            var text = DataReader.ReadText(TextTag.Expression);
            var textElement = DocxHelper.CreateTextElement(element, element.Parent, text);
            XElement result = null;
            switch (this.DynamicContentMode)
            {
                case DynamicContentMode.NoLock:
                    result = textElement;
                    break;
                case DynamicContentMode.Lock:
                    result = DocxHelper.CreateDynamicContentElement(new[] { textElement }, TextTag.TagNode);
                    break;
            }
            element.AddBeforeSelf(result);
            element.Remove();
        }
    }
}