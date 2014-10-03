using System.Collections.Generic;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    internal abstract class AbstractProcessor : ITagProcessor
    {
        private ICollection<ITagProcessor> processors = new List<ITagProcessor>();

        public DataReader DataReader { get; set; }

        public virtual void Process()
        {
            foreach (var processor in processors)
            {
                processor.DataReader = DataReader;
                processor.Process();
            }
        }

        public void AddProcessor(ITagProcessor processor)
        {
            processors.Add(processor);
        }
    }
}