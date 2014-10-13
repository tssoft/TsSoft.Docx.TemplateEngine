using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;

    internal class DocxHelper
    {
        private static readonly HashSet<XName> ValidTextRunContainers = new HashSet<XName>
                                                                                {
                                                                                    WordMl.ParagraphName,
                                                                                    WordMl.CustomXmlName,
                                                                                    WordMl.HyperlinkName,
                                                                                    WordMl.MoveFromName,
                                                                                    WordMl.MoveToName,
                                                                                    WordMl.SdtContentName,
                                                                                    WordMl.DeleteName,
                                                                                    WordMl.InsertName,
                                                                                    WordMl.SmartTagName,
                                                                                    WordMl.FieldSimpleName
                                                                                };

        private static readonly HashSet<XName> TextPropertiesNames = new HashSet<XName>
                                                                                {
                                                                                    TextRunProperties.BoldName,
                                                                                    TextRunProperties.RunFontsName,
                                                                                    TextRunProperties.ReferenceStyleName,
                                                                                    TextRunProperties.ComplexScriptBoldName,
                                                                                    TextRunProperties.ItalicsName,
                                                                                    TextRunProperties.CapsName,
                                                                                    TextRunProperties.SmallCapsName,
                                                                                    TextRunProperties.StrikeName,
                                                                                    TextRunProperties.DoubleStrikeName,
                                                                                    TextRunProperties.OutlineName,
                                                                                    TextRunProperties.ShadowName,
                                                                                    TextRunProperties.EmbossName,
                                                                                    TextRunProperties.ImprintName,
                                                                                    TextRunProperties.SnapToGridName,
                                                                                    TextRunProperties.VanishName,
                                                                                    TextRunProperties.WebHiddenName,
                                                                                    TextRunProperties.ColorName,
                                                                                    TextRunProperties.SpacingName,
                                                                                    TextRunProperties.WName,
                                                                                    TextRunProperties.KernName,
                                                                                    TextRunProperties.PositionName,
                                                                                    TextRunProperties.FontSizeName,
                                                                                    TextRunProperties.ComplexScriptFontSizeName,
                                                                                    TextRunProperties.HighlightTextName,
                                                                                    TextRunProperties.UnderlineName,
                                                                                    TextRunProperties.EffectName,
                                                                                    TextRunProperties.TextBorderName,
                                                                                    TextRunProperties.RunShadingName,
                                                                                    TextRunProperties.FitTextName,
                                                                                    TextRunProperties.VertAlignName,
                                                                                    TextRunProperties.RightToLeftTextName,
                                                                                    TextRunProperties.ComplexScriptName,
                                                                                    TextRunProperties.EmphasisMarkName,
                                                                                    TextRunProperties.EastAsianLayoutName,
                                                                                    TextRunProperties.SpecVanishName,
                                                                                    TextRunProperties.OfficeOpenXmlMath,
                                                                                };

        private static readonly HashSet<XName> CellPropertiesNames = new HashSet<XName>
                                                                         {
                                                                             TableCellProperties.BorderCellName,
                                                                             TableCellProperties.ConditionFormattingCellName,
                                                                             TableCellProperties.DeletionCellName,
                                                                             TableCellProperties.FitTextCellName,
                                                                             TableCellProperties.GridSpanCellName,
                                                                             TableCellProperties.HorizontallyMergedCellName,
                                                                             TableCellProperties.IgnoreEndMarkerCellName,
                                                                             TableCellProperties.InsertionCellName,
                                                                             TableCellProperties.NoWrapCellName,
                                                                             TableCellProperties.PreferredWidthCellName,
                                                                             TableCellProperties.RevisionPropertiesInformationCellName,
                                                                             TableCellProperties.ShadingCellName,
                                                                             TableCellProperties.SingleMarginsCellName,
                                                                             TableCellProperties.TextFlowDirectionCellName,
                                                                             TableCellProperties.VerticalAlignmentCellName,
                                                                             TableCellProperties.VerticallyMergedCellName,
                                                                             TableCellProperties.VerticallyMergedSplitCellName
                                                                         };

        private static readonly HashSet<XName> ParagraphPropertiesNames = new HashSet<XName>
                                                                              {
                                                                                  ParagraphProperties.ReferenceStyleName,
                                                                                  ParagraphProperties.KeepNextName,
                                                                                  ParagraphProperties.KeepLineName,
                                                                                  ParagraphProperties.PageBreakName,
                                                                                  ParagraphProperties.TextFrameName,
                                                                                  ParagraphProperties.WidowControlName,
                                                                                  ParagraphProperties.NumberingDefinitionName,
                                                                                  ParagraphProperties.SuppressLineNumbersName,
                                                                                  ParagraphProperties.BordersName,
                                                                                  ParagraphProperties.ShadingName,
                                                                                  ParagraphProperties.CustomTabName,
                                                                                  ParagraphProperties.SuppressHyphenationName,
                                                                                  ParagraphProperties.KinsokuName,
                                                                                  ParagraphProperties.WordWrapName,
                                                                                  ParagraphProperties.OverflowPunctuationName,
                                                                                  ParagraphProperties.TopLinePunctuationName,
                                                                                  ParagraphProperties.AutoSpaceDEName,
                                                                                  ParagraphProperties.AutoSpaceDNName,
                                                                                  ParagraphProperties.RightToLeftLayoutName,
                                                                                  ParagraphProperties.AdjuctRightIndentName,
                                                                                  ParagraphProperties.SnapToGridName,
                                                                                  ParagraphProperties.SpacingBetweenLineName,
                                                                                  ParagraphProperties.IndentationName,
                                                                                  ParagraphProperties.ContextualSpacingName,
                                                                                  ParagraphProperties.MirrorIndentsName,
                                                                                  ParagraphProperties.SuppressOverlapName,
                                                                                  ParagraphProperties.AlignmentName,
                                                                                  ParagraphProperties.TextDirectionName,
                                                                                  ParagraphProperties.TextAlignmentName,
                                                                                  ParagraphProperties.TextboxTightWrapName,
                                                                                  ParagraphProperties.OutlineLevelName,
                                                                                  ParagraphProperties.DivIdName,
                                                                                  ParagraphProperties.ConditionalFormattingStyleName,
                                                                                  ParagraphProperties.RunPropertiesName,
                                                                                  ParagraphProperties.SectionPropertiesName,
                                                                                  ParagraphProperties.RevisionPropertiesInformationName
                                                                              };

        public static XElement CreateTextElement(XElement self, XElement parent, string text)
        {
            return CreateTextElement(self, parent, text, new XElement(WordMl.ParagraphName));
        }

        public static XElement CreateTextElement(XElement self, XElement parent, string text, XName wrapTo)
        {
            return CreateTextElement(self, parent, text, new XElement(wrapTo));
        }

        public static XElement CreateTextElement(XElement self, XElement parent, string text, XElement wrapTo)
        {
            var result = new XElement(WordMl.TextRunName, new XElement(WordMl.TextName, text));
            IEnumerable<XElement> paragraphProperties = null;
            if (self.IsSdt())
            {
                var properties = self.Element(WordMl.SdtContentName)
                    .Descendants(WordMl.TextRunPropertiesName)
                    .Elements()
                    .Where(e => TextPropertiesNames.Contains(e.Name))
                    .Distinct(new NameComparer());
                result.AddFirst(new XElement(WordMl.TextRunPropertiesName, properties));
                if (self.Element(WordMl.SdtContentName).Elements(WordMl.ParagraphName).Any())
                {
                    paragraphProperties =
                        self.Element(WordMl.SdtContentName)
                            .Descendants(WordMl.ParagraphPropertiesName)
                            .Elements()
                            .Where(e => ParagraphPropertiesNames.Contains(e.Name))
                            .Distinct(new NameComparer());
                }
            }
            if (paragraphProperties == null && parent.Elements(WordMl.ParagraphName).Any())
            {
                paragraphProperties =
                    parent.Descendants(WordMl.ParagraphPropertiesName)
                          .Elements()
                          .Where(e => ParagraphPropertiesNames.Contains(e.Name))
                          .Distinct(new NameComparer());
            }
            if (paragraphProperties != null)
            {
                wrapTo.AddFirst(new XElement(WordMl.ParagraphPropertiesName, paragraphProperties));
            }
            if (!ValidTextRunContainers.Any(name => name.Equals(parent.Name)))
            {
                wrapTo.Add(result);
                result = wrapTo;
            }
            return result;
        }

        private class NameComparer : IEqualityComparer<XElement>
        {
            public bool Equals(XElement x, XElement y)
            {
                return x.Name.Equals(y.Name);
            }

            public int GetHashCode(XElement obj)
            {
                return obj.Name.GetHashCode();
            }
        }
    }
}