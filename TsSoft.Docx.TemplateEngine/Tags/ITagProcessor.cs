namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal abstract class ITagProcessor<E>
    {
        public DataReader<E> DataReader { get; set; }
    }
}