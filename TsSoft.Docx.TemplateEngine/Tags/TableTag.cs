using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    /// <summary>
    /// Table properties
    /// </summary>
    /// <author>Георгий Поликарпов</author>
    class TableTag
    {
        public string ItemsSource { get; set; }
        public int DynamicRow { get; set; }
        public XElement Table { get; set; }
    }
}
