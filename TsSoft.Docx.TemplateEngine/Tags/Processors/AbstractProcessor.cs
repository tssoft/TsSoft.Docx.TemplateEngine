using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    using System.Diagnostics.CodeAnalysis;

    internal abstract class AbstractProcessor : ITagProcessor
    {
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1304:NonPrivateReadonlyFieldsMustBeginWithUpperCaseLetter", Justification = "Reviewed. Suppression is OK here.")]
        protected readonly ICollection<ITagProcessor> processors = new List<ITagProcessor>();

        public ICollection<ITagProcessor> Processors
        {
            get
            {
                return this.processors;
            }
        }

        public virtual DataReader DataReader { get; set; }

        public bool CreateDynamicContentTags { get; set; }

        public SdtTagLockingType DynamicContentLockingType { get; set; }

        public virtual void Process()
        {
            foreach (var processor in this.processors)
            {
                processor.DataReader = DataReader;
                processor.CreateDynamicContentTags = this.CreateDynamicContentTags;
                processor.DynamicContentLockingType = this.DynamicContentLockingType;
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
            TraverseUtils.ElementsBetween(from, to).Remove();
            from.Remove();
            to.Remove();
        }
    }
}