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
            for (int listIndex = 0; listIndex < tableElementsList.Count; listIndex++)
            {
                var currentTableElement = tableElementsList[listIndex];
                if (currentTableElement.IsItemIf)
                {
                    previous = this.ProcessItemIfElement(currentTableElement, dataReader, index, previous);
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

        private XElement ProcessItemIfElement(TableElement itemIfElement, DataReader dataReader, int index, XElement previous)
        {
            bool condition;
            bool.TryParse(dataReader.ReadText(itemIfElement.Expression), out condition);
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

        private void RemoveTags()
        {
            this.CleanUp(TableTag.TagTable, TableTag.TagContent);
            this.CleanUp(TableTag.TagEndContent, TableTag.TagEndTable);
        }
               

    }
}
