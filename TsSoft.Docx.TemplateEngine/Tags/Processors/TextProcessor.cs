namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{

    internal class TextProcessor : AbstractProcessor
    {
        public TextTag TextTag { get; set; }

        public override void Process()
        {
            base.Process();
            var element = TextTag.TagNode;
            var text = DataReader.ReadText(TextTag.Expression);
            var textElement = DocxHelper.CreateTextElement(element, element.Parent, text);
            var result = this.LockDynamicContent
                             ? DocxHelper.CreateDynamicContentElement(new[] { textElement }, this.TextTag.TagNode)
                             : textElement;
            element.AddBeforeSelf(result);
            element.Remove();
        }
    }
}