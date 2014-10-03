using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    /// <summary>
    /// Table
    /// </summary>
    internal class TableTag
    {
        public int? DynamicRow { get; set; }

        public string ItemsSource { get; set; }

        public XElement Table { get; set; }

        public XElement TagTable { get; set; }

        public XElement TagContent { get; set; }

        public XElement TagEndContent { get; set; }

        public XElement TagEndTable { get; set; }
    }
}