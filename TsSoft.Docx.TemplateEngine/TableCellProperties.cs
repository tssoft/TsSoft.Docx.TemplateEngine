using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine
{
    internal static class TableCellProperties
    {
        public static readonly XName BorderCellName = WordMl.WordMlNamespace + "tcBorders";

        public static readonly XName ConditionFormattingCellName = WordMl.WordMlNamespace + "cnfStyle";

        public static readonly XName DeletionCellName = WordMl.WordMlNamespace + "cellDel";

        public static readonly XName FitTextCellName = WordMl.WordMlNamespace + "tcFitText";

        public static readonly XName GridSpanCellName = WordMl.WordMlNamespace + "gridSpan";

        public static readonly XName HorizontallyMergedCellName = WordMl.WordMlNamespace + "hMerge";

        public static readonly XName IgnoreEndMarkerCellName = WordMl.WordMlNamespace + "hideMark";

        public static readonly XName InsertionCellName = WordMl.WordMlNamespace + "cellIns";

        public static readonly XName NoWrapCellName = WordMl.WordMlNamespace + "noWrap";

        public static readonly XName PreferredWidthCellName = WordMl.WordMlNamespace + "tcW";

        public static readonly XName RevisionPropertiesInformationCellName = WordMl.WordMlNamespace + "tcPrChange";

        public static readonly XName ShadingCellName = WordMl.WordMlNamespace + "shd";

        public static readonly XName SingleMarginsCellName = WordMl.WordMlNamespace + "tcMar";

        public static readonly XName TextFlowDirectionCellName = WordMl.WordMlNamespace + "textDirection";

        public static readonly XName VerticalAlignmentCellName = WordMl.WordMlNamespace + "vAlign";

        public static readonly XName VerticallyMergedCellName = WordMl.WordMlNamespace + "vMerge";

        public static readonly XName VerticallyMergedSplitCellName = WordMl.WordMlNamespace + "cellMerge";
    }
}
