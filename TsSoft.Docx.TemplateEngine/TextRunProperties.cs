using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine
{
    internal static class TextRunProperties
    {
        public static readonly XName ReferenceStyleName = WordMl.WordMlNamespace + "rStyle";
        public static readonly XName RunFontsName = WordMl.WordMlNamespace + "rFonts";
        public static readonly XName BoldName = WordMl.WordMlNamespace + "b";
        public static readonly XName ComplexScriptBoldName = WordMl.WordMlNamespace + "bCs";
        public static readonly XName ItalicsName = WordMl.WordMlNamespace + "i";
        public static readonly XName ComplexScriptItalicsName = WordMl.WordMlNamespace + "iCs";
        public static readonly XName CapsName = WordMl.WordMlNamespace + "caps";
        public static readonly XName SmallCapsName = WordMl.WordMlNamespace + "smallCaps";
        public static readonly XName StrikeName = WordMl.WordMlNamespace + "strike";
        public static readonly XName DoubleStrikeName = WordMl.WordMlNamespace + "dstrike";
        public static readonly XName OutlineName = WordMl.WordMlNamespace + "outline";
        public static readonly XName ShadowName = WordMl.WordMlNamespace + "shadow";
        public static readonly XName EmbossName = WordMl.WordMlNamespace + "emboss";
        public static readonly XName ImprintName = WordMl.WordMlNamespace + "imprint";
        public static readonly XName NoProofName = WordMl.WordMlNamespace + "noProof";
        public static readonly XName SnapToGridName = WordMl.WordMlNamespace + "snapToGrid";
        public static readonly XName VanishName = WordMl.WordMlNamespace + "vanish";
        public static readonly XName WebHiddenName = WordMl.WordMlNamespace + "webHidden";
        public static readonly XName ColorName = WordMl.WordMlNamespace + "color";
        public static readonly XName SpacingName = WordMl.WordMlNamespace + "spacing";
        public static readonly XName WName = WordMl.WordMlNamespace + "w";
        public static readonly XName KernName = WordMl.WordMlNamespace + "kern";
        public static readonly XName PositionName = WordMl.WordMlNamespace + "position";
        public static readonly XName FontSizeName = WordMl.WordMlNamespace + "sz";
        public static readonly XName ComplexScriptFontSizeName = WordMl.WordMlNamespace + "szCs";
        public static readonly XName HighlightTextName = WordMl.WordMlNamespace + "highlight";
        public static readonly XName UnderlineName = WordMl.WordMlNamespace + "u";
        public static readonly XName EffectName = WordMl.WordMlNamespace + "effect";
        public static readonly XName TextBorderName = WordMl.WordMlNamespace + "bdr";
        public static readonly XName RunShadingName = WordMl.WordMlNamespace + "shd";
        public static readonly XName FitTextName = WordMl.WordMlNamespace + "fitText";
        public static readonly XName VertAlignName = WordMl.WordMlNamespace + "vertAlign";
        public static readonly XName RightToLeftTextName = WordMl.WordMlNamespace + "rtl";
        public static readonly XName ComplexScriptName = WordMl.WordMlNamespace + "cs";
        public static readonly XName EmphasisMarkName = WordMl.WordMlNamespace + "em";
        public static readonly XName LanguageName = WordMl.WordMlNamespace + "lang";
        public static readonly XName EastAsianLayoutName = WordMl.WordMlNamespace + "eastAsianLayout";
        public static readonly XName SpecVanishName = WordMl.WordMlNamespace + "specVanish";
        public static readonly XName OfficeOpenXmlMath = WordMl.WordMlNamespace + "oMath";
    }
}
