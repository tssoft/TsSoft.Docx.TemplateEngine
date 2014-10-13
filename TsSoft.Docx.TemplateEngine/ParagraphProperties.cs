using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine
{
    internal static class ParagraphProperties
    {
        public static readonly XName ReferenceStyleName = WordMl.WordMlNamespace + "pStyle";

        public static readonly XName KeepNextName = WordMl.WordMlNamespace + "keepNext";

        public static readonly XName KeepLineName = WordMl.WordMlNamespace + "keepLines";

        public static readonly XName PageBreakName = WordMl.WordMlNamespace + "pageBreakBefore";

        public static readonly XName TextFrameName = WordMl.WordMlNamespace + "framePr";

        public static readonly XName WidowControlName = WordMl.WordMlNamespace + "widowControl";

        public static readonly XName NumberingDefinitionName = WordMl.WordMlNamespace + "numPr";

        public static readonly XName SuppressLineNumbersName = WordMl.WordMlNamespace + "suppressLineNumbers";

        public static readonly XName BordersName = WordMl.WordMlNamespace + "pBdr";

        public static readonly XName ShadingName = WordMl.WordMlNamespace + "shd";

        public static readonly XName CustomTabName = WordMl.WordMlNamespace + "tabs";

        public static readonly XName SuppressHyphenationName = WordMl.WordMlNamespace + "suppressAutoHyphens";

        public static readonly XName KinsokuName = WordMl.WordMlNamespace + "kinsoku";

        public static readonly XName WordWrapName = WordMl.WordMlNamespace + "wordWrap";

        public static readonly XName OverflowPunctuationName = WordMl.WordMlNamespace + "overflowPunct";

        public static readonly XName TopLinePunctuationName = WordMl.WordMlNamespace + "topLinePunct";

        public static readonly XName AutoSpaceDEName = WordMl.WordMlNamespace + "autoSpaceDE";

        public static readonly XName AutoSpaceDNName = WordMl.WordMlNamespace + "autoSpaceDN";

        public static readonly XName RightToLeftLayoutName = WordMl.WordMlNamespace + "bidi";

        public static readonly XName AdjuctRightIndentName = WordMl.WordMlNamespace + "adjustRightInd";

        public static readonly XName SnapToGridName = WordMl.WordMlNamespace + "snapToGrid";

        public static readonly XName SpacingBetweenLineName = WordMl.WordMlNamespace + "spacing";

        public static readonly XName IndentationName = WordMl.WordMlNamespace + "ind";

        public static readonly XName ContextualSpacingName = WordMl.WordMlNamespace + "contextualSpacing";

        public static readonly XName MirrorIndentsName = WordMl.WordMlNamespace + "mirrorIndents";

        public static readonly XName SuppressOverlapName = WordMl.WordMlNamespace + "suppressOverlap";

        public static readonly XName AlignmentName = WordMl.WordMlNamespace + "jc";

        public static readonly XName TextDirectionName = WordMl.WordMlNamespace + "textDirection";

        public static readonly XName TextAlignmentName = WordMl.WordMlNamespace + "textAlignment";

        public static readonly XName TextboxTightWrapName = WordMl.WordMlNamespace + "textboxTightWrap";

        public static readonly XName OutlineLevelName = WordMl.WordMlNamespace + "outlineLvl";

        public static readonly XName DivIdName = WordMl.WordMlNamespace + "divId";

        public static readonly XName ConditionalFormattingStyleName = WordMl.WordMlNamespace + "cnfStyle";

        public static readonly XName RunPropertiesName = WordMl.WordMlNamespace + "rPr";

        public static readonly XName SectionPropertiesName = WordMl.WordMlNamespace + "sectPr";

        public static readonly XName RevisionPropertiesInformationName = WordMl.WordMlNamespace + "pPrChange";
    }
}
