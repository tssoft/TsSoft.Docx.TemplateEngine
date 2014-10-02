namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    internal interface ITagProcessor
    {
        DataReader DataReader { get; set; }

        void Process();

        void AddProcessor(ITagProcessor processor);
    }
}