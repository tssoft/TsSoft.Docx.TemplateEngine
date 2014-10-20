using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Table
    /// </summary>
    internal class TableTag
    {
        public int? DynamicRow { get; set; }

        public string ItemsSource { get; set; }

        public XElement Table { get; set; }

        public IEnumerable<TableElement> TagElements { get; set; }

        public XElement TagTable { get; set; }

        public XElement TagEndTable { get; set; }

        public Func<XElement, IEnumerable<TableElement>> MakeTableElementCallback { get; set; }
    }
}