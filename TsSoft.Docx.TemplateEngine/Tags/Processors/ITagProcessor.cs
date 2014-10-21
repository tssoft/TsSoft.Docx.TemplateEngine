namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    using System.Collections.Generic;

    internal interface ITagProcessor
    {
        DataReader DataReader { get; set; }

        DynamicContentMode DynamicContentMode { get; set; }

        ICollection<ITagProcessor> Processors { get; }

        void Process();

        void AddProcessor(ITagProcessor processor);
    }
}