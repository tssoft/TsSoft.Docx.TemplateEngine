namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    using System.Collections.Generic;

    internal interface ITagProcessor
    {
        DataReader DataReader { get; set; }

        void Process();

        void AddProcessor(ITagProcessor processor);

        ICollection<ITagProcessor> Processors { get; }
    }
}