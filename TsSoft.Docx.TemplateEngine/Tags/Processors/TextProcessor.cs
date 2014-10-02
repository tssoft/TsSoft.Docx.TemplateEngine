namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    internal class TextProcessor : AbstractProcessor
    {
        public TextTag TextTag { get; set; }

        public override void Process()
        {
            var element = TextTag.TagNode;
            var text = DataReader.ReadText(TextTag.Expression);
            element.AddBeforeSelf(DocxHelper.CreateTextElement(text));
            element.Remove();
        }
    }
}