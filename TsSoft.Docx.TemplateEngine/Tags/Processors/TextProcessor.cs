namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    internal class TextProcessor : ITagProcessor
    {
        public TextTag TextTag { get; set; }

        public TextProcessor(TextTag textTag)
        {
            this.TextTag = textTag;
        }

        public void Process()
        {
            var element = TextTag.TagNode;
            var text = DataReader.ReadText(TextTag.Expression);
            element.AddBeforeSelf(DocxHelper.CreateTextElement(text));
            element.Remove();
        }
    }
}