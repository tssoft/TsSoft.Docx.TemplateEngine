using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    /// <summary>
    /// Table
    /// </summary>
    internal class TableTag
    {
        public string ItemsSource { get; set; }

        public int DynamicRow { get; set; }

        public XElement Table { get; set; }
    }
}