using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine
{
    internal static class WordMl
    {
        public static readonly XNamespace WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XName BodyName = WordMlNamespace + "body";
        public static readonly XName IdName = WordMlNamespace + "id";
        public static readonly XName ParagraphName = WordMlNamespace + "p";
        public static readonly XName ParagraphPropertiesName = WordMlNamespace + "pPr";
        public static readonly XName SdtContentName = WordMlNamespace + "sdtContent";
        public static readonly XName SdtName = WordMlNamespace + "sdt";
        public static readonly XName SdtPrName = WordMlNamespace + "sdtPr";
        public static readonly XName TableName = WordMlNamespace + "tbl";
        public static readonly XName TableCellName = WordMlNamespace + "tc";
        public static readonly XName TableRowName = WordMlNamespace + "tr";
        public static readonly XName TagName = WordMlNamespace + "tag";
        public static readonly XName ValAttributeName = WordMlNamespace + "val";
        public static readonly XName TextRunName = WordMlNamespace + "r";
        public static readonly XName TextRunPropertiesName = WordMlNamespace + "rPr";
        public static readonly XName TextName = WordMlNamespace + "t";
        public static readonly XName CustomXmlName = WordMlNamespace + "customXml";
        public static readonly XName HyperlinkName = WordMlNamespace + "hyperlink";
        public static readonly XName FieldSimpleName = WordMlNamespace + "fldSimple";
        public static readonly XName SmartTagName = WordMlNamespace + "smartTag";
        public static readonly XName InsertName = WordMlNamespace + "ins";
        public static readonly XName DeleteName = WordMlNamespace + "del";
        public static readonly XName MoveFromName = WordMlNamespace + "moveFrom";
        public static readonly XName MoveToName = WordMlNamespace + "moveTo";
        public static readonly XName ProofingErrorAnchorName = WordMlNamespace + "proofErr";
        public static readonly XName BookmarkStartName = WordMlNamespace + "bookmarkStart";

        public static readonly XName ReferenceStyleName = WordMlNamespace + "rStyle";
        public static readonly XName RunFontsName = WordMlNamespace + "rFonts";
        public static readonly XName BoldName = WordMlNamespace + "b";
        public static readonly XName ComplexScriptBoldName = WordMlNamespace + "bCs";
        public static readonly XName ItalicsName = WordMlNamespace + "i";
        public static readonly XName ComplexScriptItalicsName = WordMlNamespace + "iCs";
        public static readonly XName CapsName = WordMlNamespace + "caps";
        public static readonly XName SmallCapsName = WordMlNamespace + "smallCaps";
        public static readonly XName StrikeName = WordMlNamespace + "strike";
        public static readonly XName DoubleStrikeName = WordMlNamespace + "dstrike";
        public static readonly XName OutlineName = WordMlNamespace + "outline";
        public static readonly XName ShadowName = WordMlNamespace + "shadow";
        public static readonly XName EmbossName = WordMlNamespace + "emboss";
        public static readonly XName ImprintName = WordMlNamespace + "imprint";
        public static readonly XName NoProofName = WordMlNamespace + "noProof";
        public static readonly XName SnapToGridName = WordMlNamespace + "snapToGrid";
        public static readonly XName VanishName = WordMlNamespace + "vanish";
        public static readonly XName WebHiddenName = WordMlNamespace + "webHidden";
        public static readonly XName ColorName = WordMlNamespace + "color";
        public static readonly XName SpacingName = WordMlNamespace + "spacing";
        public static readonly XName WName = WordMlNamespace + "w";
        public static readonly XName KernName = WordMlNamespace + "kern";
        public static readonly XName PositionName = WordMlNamespace + "position";
        public static readonly XName FontSizeName = WordMlNamespace + "sz";
        public static readonly XName ComplexScriptFontSizeName = WordMlNamespace + "szCs";
        public static readonly XName HighlightTextName = WordMlNamespace + "highlight";
        public static readonly XName UnderlineName = WordMlNamespace + "u";
        public static readonly XName EffectName = WordMlNamespace + "effect";
        public static readonly XName TextBorderName = WordMlNamespace + "bdr";
        public static readonly XName RunShadingName = WordMlNamespace + "shd";
        public static readonly XName FitTextName = WordMlNamespace + "fitText";
        public static readonly XName VertAlignName = WordMlNamespace + "vertAlign";
        public static readonly XName RightToLeftTextName = WordMlNamespace + "rtl";
        public static readonly XName ComplexScriptName = WordMlNamespace + "cs";
        public static readonly XName EmphasisMarkName = WordMlNamespace + "em";
        public static readonly XName LanguageName = WordMlNamespace + "lang";
        public static readonly XName EastAsianLayoutName = WordMlNamespace + "eastAsianLayout";
        public static readonly XName SpecVanishName = WordMlNamespace + "specVanish";
        public static readonly XName OfficeOpenXmlMath = WordMlNamespace + "oMath";
        


    }
}