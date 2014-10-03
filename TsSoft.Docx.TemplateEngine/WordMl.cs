using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine
{
    internal static class WordMl
    {
        public static readonly XNamespace WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XName SdtName = WordMlNamespace + "sdt";
        public static readonly XName SdtPrName = WordMlNamespace + "sdtPr";
        public static readonly XName TagName = WordMlNamespace + "tag";
        public static readonly XName SdtContentName = WordMlNamespace + "sdtContent";
        public static readonly XName ValAttributeName = WordMlNamespace + "val";
        public static readonly XName RName = WordMlNamespace + "r";
    }
}