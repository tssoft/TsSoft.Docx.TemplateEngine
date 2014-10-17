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

            if (TableTag == null)
            {
                throw new NullReferenceException();
            }

            var tableRows = TableTag.Table.Elements(WordMl.TableRowName);
            var dynamicRow = TableTag.DynamicRow.HasValue ? tableRows.ElementAt(TableTag.DynamicRow.Value - 1) : tableRows.Last();

            ReplaceValues(dynamicRow);

            RemoveTags();
        }

        private void ReplaceValues(XElement dynamicRow)
        {
            var readers = DataReader.GetReaders(TableTag.ItemsSource).ToList();
            for (int index = 0; index < readers.Count(); index++)
            {
                var currentRow = new XElement(dynamicRow);
                var tableElements = TableTag.MakeTableElementCallback(currentRow);
                this.ProcessElements(tableElements, readers[index], index+1, null);

                dynamicRow.AddBeforeSelf(currentRow);
            }
            dynamicRow.Remove();
        }

        private void ProcessElements(IEnumerable<TableElement> tableElements, DataReader dataReader, int index, XElement start)
        {
            var tableElementsList = tableElements.ToList();
            XElement previous = start;
            XElement firstCell = null;
            for (int listIndex = 0; listIndex < tableElementsList.Count; listIndex++)
            {
                string resultText = null;
                var currentTableElement = tableElementsList[listIndex];
                if (currentTableElement.IsItem)
                {
                    resultText = dataReader.ReadText(currentTableElement.Expression);
                }
                else if (currentTableElement.IsItemIf)
                {
                    bool condition;
                    bool.TryParse(dataReader.ReadText(currentTableElement.Expression), out condition);
                    if (condition)
                    {
                        this.ProcessElements(currentTableElement.TagElements, dataReader, index, previous);
                        previous = currentTableElement.StartTag.Ancestors().First(element => element.Name == WordMl.TableCellName);
                        currentTableElement.StartTag.Remove();
                        currentTableElement.EndTag.Remove();
                    }
                    else
                    {
                        var currentCell = currentTableElement.StartTag.Ancestors().First(element => element.Name == WordMl.TableCellName);
                        if (previous != null && !previous.Equals(currentCell))
                        {
                            currentCell.Remove();
                            previous.AddAfterSelf(currentCell);
                        }
                        previous = currentCell;
                        this.CleanUp(currentTableElement.StartTag, currentTableElement.EndTag);
                    }
                }
                else if (currentTableElement.IsIndex)
                {
                    resultText = index.ToString(CultureInfo.CurrentCulture);
                }
                if (resultText != null)
                {
                    previous = this.ProcessCell(currentTableElement, previous, resultText);
                    if (firstCell == null)
                    {
                        firstCell = previous;
                    }
                    
                    currentTableElement.StartTag.Remove();
                }
            }

            if (start == null && firstCell != null)
            {
                var staticCells = firstCell.ElementsBeforeSelf(WordMl.TableCellName).ToList();
                if (staticCells.Any())
                {
                    staticCells.Remove();
                    previous.AddAfterSelf(staticCells);
                }
            }
        }

        private XElement ProcessCell(TableElement tableElement, XElement previous, string text)
        {
            var isInnerCell = true;
            var currentCell = tableElement.StartTag.Descendants(WordMl.TableCellName).FirstOrDefault();
            if (currentCell == null)
            {
                currentCell = tableElement.StartTag.Ancestors().First(element => element.Name == WordMl.TableCellName);
                isInnerCell = false;
            }

            XElement parent = null;
            if (isInnerCell)
            {
                parent = currentCell;
            }
            else
            {
                parent = currentCell.Elements(WordMl.ParagraphName).Any() ? currentCell.Element(WordMl.ParagraphName) : tableElement.StartTag;
            }
            var result = DocxHelper.CreateTextElement(
                tableElement.StartTag,
                parent,
                text);

            var staticCells = isInnerCell
                                  ? tableElement.StartTag.ElementsBeforeSelf(WordMl.TableCellName).ToList()
                                  : currentCell.ElementsBeforeSelf(WordMl.TableCellName).ToList();
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

            if (!isInnerCell)
            {
                tableElement.StartTag.AddAfterSelf(result);
                if (previous == null)
                {
                    var parentRow = currentCell.Parent;
                    currentCell.Remove();
                    parentRow.Add(currentCell);
                }
            }
            else
            {
                if (currentCell.Elements(WordMl.ParagraphName).Any())
                {
                    currentCell.Elements(WordMl.ParagraphName).Remove();
                }
                currentCell.Add(result);

                if (previous == null)
                {
                    currentCell.Remove();
                    tableElement.StartTag.Parent.Add(currentCell);
                }
            }

            if (previous != null && !previous.Equals(currentCell))
            {
                currentCell.Remove();
                previous.AddAfterSelf(currentCell);
            }

            return currentCell;
        }

        private void RemoveTags()
        {
            this.CleanUp(TableTag.TagTable, TableTag.TagContent);
            this.CleanUp(TableTag.TagEndContent, TableTag.TagEndTable);
        }
               

    }
}
