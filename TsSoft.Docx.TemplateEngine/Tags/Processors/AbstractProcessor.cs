using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    internal abstract class AbstractProcessor : ITagProcessor
    {
        private ICollection<ITagProcessor> processors = new List<ITagProcessor>();

        public virtual DataReader DataReader { get; set; }

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

        /// <summary>
        /// Removes all elements between and including passed elements
        /// </summary>
        /// <param name="from"></param>
        /// <param name="to"></param>
        public void CleanUp(XElement from, XElement to)
        {
            foreach (var element in from.ElementsAfterSelf().Where(e => e.IsBefore(to)))
            {
                element.Remove();
            }
            from.Remove();
            to.Remove();
        }
    }
}