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
                this.ProcessElements(tableElements, readers[index], index, null);

                dynamicRow.AddBeforeSelf(currentRow);
            }
            dynamicRow.Remove();
        }

        private void ProcessElements(IEnumerable<TableElement> tableElements, DataReader dataReader, int index, XElement start)
        {
            var tableElementsList = tableElements.ToList();
            XElement previous = start;
            for (int listIndex = 0; listIndex < tableElementsList.Count; listIndex++)
            {
                XElement result = null;
                string resultText = string.Empty;
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
                        currentCell.Remove();
                        previous.AddAfterSelf(currentCell);
                        previous = currentCell;
                        this.CleanUp(currentTableElement.StartTag, currentTableElement.EndTag);
                    }
                }
                else if (currentTableElement.IsIndex)
                {
                    resultText = index.ToString(CultureInfo.CurrentCulture);
                }
                if (!string.IsNullOrEmpty(resultText))
                {
                    result = DocxHelper.CreateTextElement(
                        currentTableElement.StartTag,
                        currentTableElement.StartTag,
                        resultText);
                    var currentCell = currentTableElement.StartTag.Descendants(WordMl.TableCellName).FirstOrDefault();
                    if (currentCell == null)
                    {
                        currentCell = currentTableElement.StartTag.Ancestors().First(element => element.Name == WordMl.TableCellName);
                        if (previous == null)
                        {
                            previous = currentCell;
                        }
                        else
                        {
                            currentCell.Remove();
                            previous.AddAfterSelf(currentCell);
                        }
                    }
                    else
                    {
                        currentCell.Remove();
                        if (previous == null)
                        {
                            currentTableElement.StartTag.Parent.Add(currentCell);
                            previous = currentCell;
                        }
                        else
                        {
                            previous.AddAfterSelf(currentCell);
                            previous = currentCell;
                        }
                        currentCell.Descendants(WordMl.ParagraphName).Remove();
                    }

                    currentCell.Add(result);
                    currentTableElement.StartTag.Remove();
                }
            }
        }

        private void RemoveTags()
        {
            this.CleanUp(TableTag.TagTable, TableTag.TagContent);
            this.CleanUp(TableTag.TagEndContent, TableTag.TagEndTable);
        }
    }
}
