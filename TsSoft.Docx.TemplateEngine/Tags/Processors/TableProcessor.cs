﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    using System.Globalization;

    internal class TableProcessor : AbstractProcessor
    {
        public TableTag TableTag { get; set; }

        public override void Process()
        {
            base.Process();
            ProcessDynamicContent();

            if (TableTag == null)
            {
                throw new NullReferenceException();
            }

            var tableRows = TableTag.Table.Elements(WordMl.TableRowName);
            var dynamicRow = TableTag.DynamicRow.HasValue ? tableRows.ElementAt(TableTag.DynamicRow.Value - 1) : tableRows.Last();

            ReplaceValues(dynamicRow);

            if (this.LockDynamicContent)
            {
                var innerElements = TraverseUtils.ElementsBetween(this.TableTag.TagTable, this.TableTag.TagEndTable).ToList();
                innerElements.Remove();
                this.TableTag.TagTable.AddBeforeSelf(DocxHelper.CreateDynamicContentElement(innerElements, this.TableTag.TagTable));
                this.CleanUp(this.TableTag.TagTable, this.TableTag.TagEndTable);
            }
            else
            {
                RemoveTags();
            }
        }

        private void ReplaceValues(XElement dynamicRow)
        {
            var readers = DataReader.GetReaders(TableTag.ItemsSource).ToList();
            for (int index = 0; index < readers.Count(); index++)
            {
                var currentRow = new XElement(dynamicRow);
                var tableElements = TableTag.MakeTableElementCallback(currentRow);

                this.ProcessElements(tableElements, readers[index], index+1, null, true);

                dynamicRow.AddBeforeSelf(currentRow);
            }
            dynamicRow.Remove();
        }
   
        private void ProcessElements(IEnumerable<TableElement> tableElements, DataReader dataReader, int index, XElement start, bool isTopLevel)
        {
            var tableElementsList = tableElements.ToList();
            XElement previous = start;
            XElement firstCell = null;
            foreach (var currentTableElement in tableElementsList)
            {
                if (currentTableElement.IsItemTable)
                {
                    var itemTableGenerator = new ItemTableGenerator();
                    itemTableGenerator.Generate(currentTableElement.StartTag, currentTableElement.EndTag, dataReader);
                    
                }
                else if (currentTableElement.IsItemHtmlContent)
                {                    
                    currentTableElement.StartTag = HtmlContentProcessor.MakeHtmlContentProcessed(currentTableElement.StartTag,
                                                                                                 dataReader.ReadText(
                                                                                                     currentTableElement.Expression));                    
                }
                else if (currentTableElement.IsItemIf)
                {
                    previous = this.ProcessItemIfElement(currentTableElement, dataReader, index, previous);

                    //TODO: this if need testing
                    if (firstCell == null)
                    {
                        firstCell = previous;
                    }
                }
                else if (currentTableElement.IsItemRepeater)
                {                                       
                    this.ProcessItemRepeaterElement(currentTableElement, dataReader, index, previous);

                }
                else if (currentTableElement.IsItem || currentTableElement.IsIndex)
                {
                    var resultText = currentTableElement.IsIndex
                                         ? index.ToString(CultureInfo.CurrentCulture)
                                         : dataReader.ReadText(currentTableElement.Expression);

                    previous = this.ProcessCell(currentTableElement, previous, resultText);
                    if (firstCell == null)
                    {
                        firstCell = previous;
                    }

                    currentTableElement.StartTag.Remove();
                }
            }

            if (isTopLevel && firstCell != null)
            {
                ProcessStaticCells(firstCell, previous);
            }
        }


        private ItemRepeaterTag GenerateItemRepeaterTag(TableElement itemRepeaterElement)
        {
            var tagResult = new ItemRepeaterTag()
            {   StartItemRepeater = itemRepeaterElement.StartTag,
                EndItemRepeater = itemRepeaterElement.EndTag,
                NestedRepeaters = new List<ItemRepeaterTag>(),
                Source = itemRepeaterElement.Expression
            };            
            return tagResult;
        }

        private void ProcessItemRepeaterElement(TableElement itemRepeaterElement, DataReader dataReader, int index,
                                                    XElement previous)
        {            
            var expression = itemRepeaterElement.Expression;
            var readers = dataReader.GetReaders(expression);
            var itemRepeaterTag = GenerateItemRepeaterTag(itemRepeaterElement);         
            var generator = new ItemRepeaterGenerator();
            generator.Generate(itemRepeaterTag, readers.ToList());
        }

        private XElement ProcessItemIfElement(TableElement itemIfElement, DataReader dataReader, int index, XElement previous)
        {
            bool condition;
            try
            {
                condition = bool.Parse(dataReader.ReadText(itemIfElement.Expression));
            }
            catch (FormatException)
            {
                condition = false;
            }
            catch (System.Xml.XPath.XPathException)
            {
                condition = false;
            }
            var currentCell = itemIfElement.StartTag.Ancestors().First(element => element.Name == WordMl.TableCellName); 
            if (condition)
            {
                this.ProcessElements(itemIfElement.TagElements, dataReader, index, previous, false);
                itemIfElement.StartTag.Remove();
                itemIfElement.EndTag.Remove();

                if (previous != null && !previous.Equals(currentCell))
                {
                    currentCell.Remove();
                    previous.AddAfterSelf(currentCell);
                }
            }
            else
            {
                this.CleanUp(itemIfElement.StartTag, itemIfElement.EndTag);
                if (!currentCell.Elements(WordMl.ParagraphName).Any())
                {
                    currentCell.Add(new XElement(WordMl.ParagraphName));
                }
                  
                if (previous != null && !previous.Equals(currentCell))
                {
                    currentCell.Remove();
                    previous.AddAfterSelf(currentCell);
                }
                else if (previous == null)
                {
                    var parentRow = currentCell.Parent;
                    currentCell.Remove();
                    parentRow.Add(currentCell);
                }
            }

            return currentCell;
        }

        private XElement ProcessCell(TableElement tableElement, XElement previous, string text)
        {
            var isInnerCell = tableElement.HasCell();
            var currentCell = this.CurrentCell(tableElement);
            previous = this.ProcessStaticCells(tableElement, previous);

            var parent = isInnerCell ? currentCell: tableElement.StartTag.Parent;
            var result = DocxHelper.CreateTextElement(
                tableElement.StartTag,
                parent,
                text);
            if (!isInnerCell)
            {
                tableElement.StartTag.AddAfterSelf(result);
            }
            else
            {
                if (currentCell.Elements(WordMl.ParagraphName).Any())
                {
                    currentCell.Elements(WordMl.ParagraphName).Remove();
                }                 
                currentCell.Add(result);
            }

            if (previous != null && !previous.Equals(currentCell))
            {
                currentCell.Remove();
                previous.AddAfterSelf(currentCell);
            }
            else if (previous == null)
            {
                var parentRow = isInnerCell ? tableElement.StartTag.Parent : currentCell.Parent;
                currentCell.Remove();
                parentRow.Add(currentCell);
            }
            
            return currentCell;
        }
        
        private XElement ProcessStaticCells(TableElement tableElement, XElement previous)
        {
            var currentElement = tableElement.HasCell()
                                  ? tableElement.StartTag
                                  : this.CurrentCell(tableElement);

            return ProcessStaticCells(currentElement, previous);
        }

        private XElement ProcessStaticCells(XElement currentElement, XElement previous)
        {
            var staticCells =
                currentElement.ElementsBeforeSelf(WordMl.TableCellName)
                              .Where(element => element.IsBefore(previous))
                              .ToList();
            if (staticCells.Any())
            {
                if (previous == null)
                {
                    var parentRow = staticCells.First().Parent;
                    staticCells.Remove();
                    parentRow.Add(staticCells);
                    previous = staticCells.Last();
                }
                else
                {
                    staticCells.Remove();
                    previous.AddAfterSelf(staticCells);
                    previous = staticCells.Last();
                }
            }

            return previous;
        }

        private XElement CurrentCell(TableElement tableElement)
        {
            return tableElement.StartTag.Descendants(WordMl.TableCellName).FirstOrDefault()
                   ?? tableElement.StartTag.Ancestors().First(element => element.Name == WordMl.TableCellName);
        }

        private void ProcessDynamicContent()
        {
            var dynamicContentTags =
                TraverseUtils.ElementsBetween(this.TableTag.TagTable, this.TableTag.TagEndTable)
                             .Where(
                                 element =>
                                 element.IsSdt()
                                 && element.Element(WordMl.SdtPrName)
                                           .Element(WordMl.TagName)
                                           .Attribute(WordMl.ValAttributeName)
                                           .Value.ToLower()
                                           .Equals("dynamiccontent"))
                             .ToList();
            foreach (var dynamicContentTag in dynamicContentTags)
            {
                var innerElements = dynamicContentTag.Element(WordMl.SdtContentName).Elements();
                dynamicContentTag.AddAfterSelf(innerElements);
                dynamicContentTag.Remove();
            }
        }

        private void RemoveTags()
        {            
            this.TableTag.TagTable.Remove();
            this.TableTag.TagEndTable.Remove();
        }
               

    }
}
