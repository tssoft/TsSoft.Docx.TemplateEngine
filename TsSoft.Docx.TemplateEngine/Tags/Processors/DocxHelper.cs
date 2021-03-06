﻿using System;
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

        public static XElement CreateTextElement(XElement self, XElement parent, string text, bool wrapParagraphs = false)
        {
            return CreateTextElement(self, parent, text, new XElement(WordMl.ParagraphName), string.Empty, wrapParagraphs);
        }

        public static XElement CreateTextElement(XElement self, XElement parent, string text, XName wrapTo, bool wrapParagraphs = false)
        {
            return CreateTextElement(self, parent, text, new XElement(wrapTo), wrapParagraphs);
        }

        public static XElement CreateTextElement(XElement self, XElement parent, string text, string styleName, bool wrapParagraphs = false)
        {
            return CreateTextElement(self, parent, text, new XElement(WordMl.ParagraphName), styleName, wrapParagraphs);
        }


        public static XElement CreateTextElement(XElement self, XElement parent, string text, XName wrapTo, string styleName, bool wrapParagraphs = false)
        {
            return CreateTextElement(self, parent, text, new XElement(wrapTo), styleName, wrapParagraphs);
        }


        public static XElement CreateTextElement(XElement self, XElement parent, string text, XElement wrapTo,
                                                 bool wrapParagraphs = false)
        {
            return CreateTextElement(self, parent, text, wrapTo, string.Empty, wrapParagraphs);
        }

        public static XElement CreateTextElement(XElement self, XElement parent, string text, XElement wrapTo, string styleName, bool wrapParagraphs = false)
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
                var textRunProperties = new XElement(WordMl.TextRunPropertiesName, properties);

                if (!(styleName.Equals(string.Empty) || properties.Any(p => p.Name.Equals(WordMl.ParagraphStyleName))))
                {
                    textRunProperties.Add(new XElement(TextRunProperties.ReferenceStyleName, new XAttribute(WordMl.ValAttributeName, styleName + "0")));
                }

                result.AddFirst(textRunProperties);
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
            if (paragraphProperties == null && (parent.Elements(WordMl.ParagraphName).Any() || (parent.Name == WordMl.ParagraphName)))
            {
                paragraphProperties =
                    parent.DescendantsAndSelf(WordMl.ParagraphPropertiesName)
                          .Elements()
                          .Where(e => ParagraphPropertiesNames.Contains(e.Name))
                          .Distinct(new NameComparer());
            }            

            if (paragraphProperties != null)
            {
                var paragraphPropertiesElement = new XElement(WordMl.ParagraphPropertiesName, paragraphProperties);

                if (!(styleName.Equals(string.Empty) || paragraphProperties.Any(p => p.Name.Equals(WordMl.ParagraphStyleName))))
                {
                    paragraphPropertiesElement.Add(new XElement(WordMl.ParagraphStyleName, new XAttribute(WordMl.ValAttributeName, styleName)));
                }

                wrapTo.AddFirst(paragraphPropertiesElement);
            }
            if ((!ValidTextRunContainers.Any(name => name.Equals(parent.Name))) || (wrapParagraphs && !parent.Elements().Any(el => el.Name.Equals(WordMl.TextRunName))))
            {
                wrapTo.Add(result);
                result = wrapTo;
            }
            return result;
        }

        public static bool IsEmptyParagraph(XElement paragraph)
        {
            if (!paragraph.Name.Equals(WordMl.ParagraphName))
            {
                throw new Exception("Element is not paragraph!");
            }
            return (paragraph.IsEmpty ||
                    ((paragraph.Elements().Count() == 1) &&
                     paragraph.Elements().Single().Name.Equals(WordMl.ParagraphPropertiesName)));
        }

        public static XElement CreateAltChunkElement(int altChunkId)
        {
            var formattedAltChunkId = string.Format("altChunkId{0}", altChunkId);
            var altChunk = new XElement(WordMl.AltChunkName, new XAttribute(RelationshipMl.IdName, formattedAltChunkId));
            return altChunk;
        }

        public static void AddEmptyParagraphInTableCell(XElement altChunkElement)
        {
            var rsidR = TraverseUtils.GenerateRandomRsidR();
            var rsidRPrAttr = altChunkElement.Ancestors(WordMl.TableRowName)
                                             .First()
                                             .Attribute(WordMl.RsidRPropertiesName);
            var rsidRAttr = altChunkElement.Ancestors(WordMl.TableRowName)
                                           .First()
                                           .Attribute(WordMl.RsidRName);
            var rsidP = (rsidRPrAttr != null) ? rsidRPrAttr.Value : rsidRAttr.Value;

            altChunkElement.AddAfterSelf(
                                         new XElement(WordMl.ParagraphName,
                                            new XAttribute(WordMl.RsidRName, rsidR),
                                            new XAttribute(WordMl.RsidRPropertiesName, rsidR),
                                            new XAttribute(WordMl.RsidRDefaultName, rsidR),
                                            new XAttribute(WordMl.RsidPName, rsidP)));
        }


        public static XElement CreateDynamicContentElement(IEnumerable<XElement> contentElements, XElement tagElement, SdtTagLockingType lockingType = SdtTagLockingType.Unlocked)
        {
            var tagId = tagElement.Element(WordMl.SdtPrName).Element(WordMl.IdName).Attribute(WordMl.ValAttributeName).Value;
            return CreateDynamicContentElement(contentElements, tagId, lockingType);
        }



        public static XElement CreateDynamicContentElement(IEnumerable<XElement> contentElements, string id, SdtTagLockingType lockingType = SdtTagLockingType.Unlocked)
        {
            var locker = new DynamicContentLocker();
            return new XElement(
                WordMl.SdtName,
                new XElement(
                    WordMl.SdtPrName,
                    new XElement(
                        WordMl.WordMlNamespace + "alias", new XAttribute(WordMl.ValAttributeName, "DynamicContent")),
                    new XElement(
                        WordMl.TagName, new XAttribute(WordMl.ValAttributeName, "DynamicContent")),
                    new XElement(WordMl.IdName, new XAttribute(WordMl.ValAttributeName, id)),
                    //new XElement(
                    //    WordMl.WordMlNamespace + "lock", new XAttribute(WordMl.ValAttributeName, "sdtContentLocked"))),
                    locker.GetLockingElement(lockingType)),
                new XElement(WordMl.SdtContentName, contentElements));
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