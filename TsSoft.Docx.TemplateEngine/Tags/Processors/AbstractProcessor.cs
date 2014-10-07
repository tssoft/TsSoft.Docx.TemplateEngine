using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    internal abstract class AbstractProcessor : ITagProcessor
    {
        protected readonly ICollection<ITagProcessor> processors = new List<ITagProcessor>();

        public ICollection<ITagProcessor> Processors
        {
            get
            {
                return this.processors;
            }
        }

        public virtual DataReader DataReader { get; set; }

        public virtual void Process()
        {
            foreach (var processor in this.processors)
            {
                processor.DataReader = DataReader;
                processor.Process();
            }
        }

        public void AddProcessor(ITagProcessor processor)
        {
            this.processors.Add(processor);
        }

        /// <summary>
        /// Removes all elements between and including passed elements
        /// </summary>
        /// <param name="from"></param>
        /// <param name="to"></param>
        public void CleanUp(XElement from, XElement to)
        {
            from.ElementsAfterSelf().Where(element => element.IsBefore(to)).Remove();
            from.Remove();
            to.Remove();
        }
    }
}