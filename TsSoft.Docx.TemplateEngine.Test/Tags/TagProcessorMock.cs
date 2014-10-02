using System;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    internal class TagProcessorMock<P> : ITagProcessor where P : ITagProcessor
    {
        public P InnerProcessor { get; private set; }

        public DataReader DataReader { get; set; }

        public void Process()
        {
            throw new NotImplementedException();
        }

        public void AddProcessor(ITagProcessor processor)
        {
            if (processor == null || !(processor is P))
            {
                throw new Exception();
            }
            InnerProcessor = (P)processor;
        }
    }
}