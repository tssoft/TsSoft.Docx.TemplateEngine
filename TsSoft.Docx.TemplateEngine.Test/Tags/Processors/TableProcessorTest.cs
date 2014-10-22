using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags.Processors
{

    [TestClass]
    public class TableProcessorTest : BaseProcessorTest
    {
        private XElement documentRoot;

        private DataReader dataReader;

        [TestInitialize]
        public void Initialize()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "TableProcessorTemplateTest.xml");
            var doc = XDocument.Load(docStream);
            this.documentRoot = doc.Root.Element(WordMl.BodyName);

            var dataStream = AssemblyResourceHelper.GetResourceStream(this, "TableProcessorDataTest.xml");
            var xmlDoc = XDocument.Load(dataStream);
            this.dataReader = DataReaderFactory.CreateReader(xmlDoc);
        }

        [TestMethod]
        public void TestProcess()
        {
            var tableTag = this.GetTableTag();

            Func<XElement, IEnumerable<TableElement>> MakeTableElementCallback = element =>
                {
                    var tags = element.Descendants(WordMl.SdtName).ToList();

                    var index =
                        tags.Find(
                            e =>
                            e.Element(WordMl.SdtPrName)
                             .Element(WordMl.TagName)
                             .Attribute(WordMl.ValAttributeName)
                             .Value.ToLower()
                             .Equals("itemindex"));
                    var subject =
                        tags.Find(
                            e =>
                            e.Element(WordMl.SdtPrName)
                             .Element(WordMl.TagName)
                             .Attribute(WordMl.ValAttributeName)
                             .Value.ToLower()
                             .Equals("itemtext") && e.Element(WordMl.SdtContentName).Value == "./Subject");
                    var issueDate =
                        tags.Find(
                            e =>
                            e.Element(WordMl.SdtPrName)
                             .Element(WordMl.TagName)
                             .Attribute(WordMl.ValAttributeName)
                             .Value.ToLower()
                             .Equals("itemtext") && e.Element(WordMl.SdtContentName).Value == "./IssueDate");
                    var expireDate =
                        tags.Find(
                            e =>
                            e.Element(WordMl.SdtPrName)
                             .Element(WordMl.TagName)
                             .Attribute(WordMl.ValAttributeName)
                             .Value.ToLower()
                             .Equals("itemtext") && e.Element(WordMl.SdtContentName).Value == "./ExpireDate");
                    var itemIf =
                        tags.Find(
                            e =>
                            e.Element(WordMl.SdtPrName)
                             .Element(WordMl.TagName)
                             .Attribute(WordMl.ValAttributeName)
                             .Value.ToLower()
                             .Equals("itemif"));
                    var endItemIf =
                        tags.Find(
                            e =>
                            e.Element(WordMl.SdtPrName)
                             .Element(WordMl.TagName)
                             .Attribute(WordMl.ValAttributeName)
                             .Value.ToLower()
                             .Equals("enditemif"));


                    IEnumerable<TableElement> tableElements = new TableElement[]
                                                                  {
                                                                      new TableElement
                                                                          {
                                                                              IsIndex = true,
                                                                              IsItem = false,
                                                                              IsItemIf = false,
                                                                              StartTag = index,
                                                                          },
                                                                      new TableElement
                                                                          {
                                                                              IsItemIf = true,
                                                                              IsIndex = false,
                                                                              IsItem = false,
                                                                              StartTag = itemIf,
                                                                              EndTag = endItemIf,
                                                                              Expression = "./HasSubject",
                                                                              TagElements =
                                                                                  new TableElement[]
                                                                                      {
                                                                                          new TableElement
                                                                                              {
                                                                                                  IsItem = true,
                                                                                                  IsIndex = false,
                                                                                                  IsItemIf = false,
                                                                                                  StartTag = subject,
                                                                                                  Expression = "./Subject"
                                                                                              }
                                                                                      }
                                                                          },
                                                                      new TableElement
                                                                          {
                                                                              IsItem = true,
                                                                              IsIndex = false,
                                                                              IsItemIf = false,
                                                                              StartTag = issueDate,
                                                                              Expression = "./IssueDate"
                                                                          },
                                                                      new TableElement
                                                                          {
                                                                              IsItem = true,
                                                                              IsIndex = false,
                                                                              IsItemIf = false,
                                                                              StartTag = expireDate,
                                                                              Expression = "./ExpireDate"
                                                                          }
                                                                  };
                    return tableElements;
                };

            tableTag.MakeTableElementCallback = MakeTableElementCallback;

            var tableProcessor = new TableProcessor
            {
                TableTag = tableTag,
                DataReader = this.dataReader
            };
            tableProcessor.Process();

            Assert.IsFalse(this.documentRoot.Elements(WordMl.SdtName).Any(element => element.IsSdt()));
            var expectedTableStructure = new[]
                {
                new[] { "#", "Certificate", "Date" },
                new[] { string.Empty, string.Empty, "Issue", "Expiration" },
                new[] { "1", "2", "3", "4" },
                new[] { "1", "Subject1", "01.04.2014", "01.10.2015" },
                new[] { "2", "Subject2", "01.03.2014", "01.09.2015" },
                new[] { "3", "Subject3", "01.01.2011", "01.01.2012" },
                new[] { "This", "row", "stays", "untouched" },
            };

            var rows = tableTag.Table.Elements(WordMl.TableRowName).ToList();
            Assert.AreEqual(expectedTableStructure.Count(), rows.Count());
            int rowIndex = 0;
            foreach (var row in rows)
            {
                var expectedRow = expectedTableStructure[rowIndex];

                var cellsInRow = row.Elements(WordMl.TableCellName).ToList();

                Assert.AreEqual(expectedRow.Count(), cellsInRow.Count());

                int cellIndex = 0;
                foreach (var cell in cellsInRow)
                {
                    Assert.AreEqual(expectedRow[cellIndex], cell.Value);
                    cellIndex++;
                }

                rowIndex++;
            }
            /*
            var tagsBetween =
                tableTag.TagTable.ElementsAfterSelf().Where(element => element.IsBefore(tableTag.TagContent));
            Assert.IsFalse(tagsBetween.Any());

            tagsBetween =
                tableTag.TagEndContent.ElementsAfterSelf().Where(element => element.IsBefore(tableTag.TagEndTable));
            Assert.IsFalse(tagsBetween.Any());*/
            this.ValidateTagsRemoved(this.documentRoot);
        }

        [TestMethod]
        public void TestProcessWithLock()
        {
            var tableTag = this.GetTableTag();

            Func<XElement, IEnumerable<TableElement>> MakeTableElementCallback = element =>
            {
                var tags = element.Descendants(WordMl.SdtName).ToList();

                var index =
                    tags.Find(
                        e =>
                        e.Element(WordMl.SdtPrName)
                         .Element(WordMl.TagName)
                         .Attribute(WordMl.ValAttributeName)
                         .Value.ToLower()
                         .Equals("itemindex"));
                var subject =
                    tags.Find(
                        e =>
                        e.Element(WordMl.SdtPrName)
                         .Element(WordMl.TagName)
                         .Attribute(WordMl.ValAttributeName)
                         .Value.ToLower()
                         .Equals("itemtext") && e.Element(WordMl.SdtContentName).Value == "./Subject");
                var issueDate =
                    tags.Find(
                        e =>
                        e.Element(WordMl.SdtPrName)
                         .Element(WordMl.TagName)
                         .Attribute(WordMl.ValAttributeName)
                         .Value.ToLower()
                         .Equals("itemtext") && e.Element(WordMl.SdtContentName).Value == "./IssueDate");
                var expireDate =
                    tags.Find(
                        e =>
                        e.Element(WordMl.SdtPrName)
                         .Element(WordMl.TagName)
                         .Attribute(WordMl.ValAttributeName)
                         .Value.ToLower()
                         .Equals("itemtext") && e.Element(WordMl.SdtContentName).Value == "./ExpireDate");
                var itemIf =
                    tags.Find(
                        e =>
                        e.Element(WordMl.SdtPrName)
                         .Element(WordMl.TagName)
                         .Attribute(WordMl.ValAttributeName)
                         .Value.ToLower()
                         .Equals("itemif"));
                var endItemIf =
                    tags.Find(
                        e =>
                        e.Element(WordMl.SdtPrName)
                         .Element(WordMl.TagName)
                         .Attribute(WordMl.ValAttributeName)
                         .Value.ToLower()
                         .Equals("enditemif"));


                IEnumerable<TableElement> tableElements = new TableElement[]
                                                                  {
                                                                      new TableElement
                                                                          {
                                                                              IsIndex = true,
                                                                              IsItem = false,
                                                                              IsItemIf = false,
                                                                              StartTag = index,
                                                                          },
                                                                      new TableElement
                                                                          {
                                                                              IsItemIf = true,
                                                                              IsIndex = false,
                                                                              IsItem = false,
                                                                              StartTag = itemIf,
                                                                              EndTag = endItemIf,
                                                                              Expression = "./HasSubject",
                                                                              TagElements =
                                                                                  new TableElement[]
                                                                                      {
                                                                                          new TableElement
                                                                                              {
                                                                                                  IsItem = true,
                                                                                                  IsIndex = false,
                                                                                                  IsItemIf = false,
                                                                                                  StartTag = subject,
                                                                                                  Expression = "./Subject"
                                                                                              }
                                                                                      }
                                                                          },
                                                                      new TableElement
                                                                          {
                                                                              IsItem = true,
                                                                              IsIndex = false,
                                                                              IsItemIf = false,
                                                                              StartTag = issueDate,
                                                                              Expression = "./IssueDate"
                                                                          },
                                                                      new TableElement
                                                                          {
                                                                              IsItem = true,
                                                                              IsIndex = false,
                                                                              IsItemIf = false,
                                                                              StartTag = expireDate,
                                                                              Expression = "./ExpireDate"
                                                                          }
                                                                  };
                return tableElements;
            };

            tableTag.MakeTableElementCallback = MakeTableElementCallback;

            var tableProcessor = new TableProcessor
            {
                TableTag = tableTag,
                DataReader = this.dataReader,
                LockDynamicContent = true
            };
            tableProcessor.Process();

            Assert.IsTrue(
                this.documentRoot.Elements(WordMl.SdtName)
                    .Any(
                        element =>
                        element.Element(WordMl.SdtPrName)
                               .Element(WordMl.TagName)
                               .Attribute(WordMl.ValAttributeName)
                               .Value.ToLower()
                               .Equals("dynamiccontent")));
            var expectedTableStructure = new[]
                {
                new[] { "#", "Certificate", "Date" },
                new[] { string.Empty, string.Empty, "Issue", "Expiration" },
                new[] { "1", "2", "3", "4" },
                new[] { "1", "Subject1", "01.04.2014", "01.10.2015" },
                new[] { "2", "Subject2", "01.03.2014", "01.09.2015" },
                new[] { "3", "Subject3", "01.01.2011", "01.01.2012" },
                new[] { "This", "row", "stays", "untouched" },
            };
            Assert.AreEqual(1, this.documentRoot.Elements(WordMl.SdtName).Count());
            var rows = tableTag.Table.Elements(WordMl.TableRowName).ToList();
            Assert.AreEqual(expectedTableStructure.Count(), rows.Count());
            int rowIndex = 0;
            foreach (var row in rows)
            {
                var expectedRow = expectedTableStructure[rowIndex];

                var cellsInRow = row.Elements(WordMl.TableCellName).ToList();

                Assert.AreEqual(expectedRow.Count(), cellsInRow.Count());

                int cellIndex = 0;
                foreach (var cell in cellsInRow)
                {
                    Assert.AreEqual(expectedRow[cellIndex], cell.Value);
                    cellIndex++;
                }

                rowIndex++;
            }
            this.ValidateTagsRemoved(this.documentRoot);
        }

        [TestMethod]
        public void TestProcessWithStaticCells()
        {
            XElement root;
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "TableProcessorWithStaticCellsTest.xml"))
            {
                var document = XDocument.Load(docStream);
                root = document.Root.Element(WordMl.BodyName);
            }
            var tableTag = this.GetTableTag(root);

            Func<XElement, IEnumerable<TableElement>> MakeTableElementCallback = element =>
            {
                var tags = element.Descendants(WordMl.SdtName).ToList();

                var index =
                    tags.Find(
                        e =>
                        e.Element(WordMl.SdtPrName)
                         .Element(WordMl.TagName)
                         .Attribute(WordMl.ValAttributeName)
                         .Value.ToLower()
                         .Equals("itemindex"));
                var issueDate =
                    tags.Find(
                        e =>
                        e.Element(WordMl.SdtPrName)
                         .Element(WordMl.TagName)
                         .Attribute(WordMl.ValAttributeName)
                         .Value.ToLower()
                         .Equals("itemtext") && e.Element(WordMl.SdtContentName).Value == "./IssueDate");


                IEnumerable<TableElement> tableElements = new TableElement[]
                                                                  {
                                                                      new TableElement
                                                                          {
                                                                              IsIndex = true,
                                                                              IsItem = false,
                                                                              IsItemIf = false,
                                                                              StartTag = index,
                                                                          },
                                                                      new TableElement
                                                                          {
                                                                              IsItem = true,
                                                                              IsIndex = false,
                                                                              IsItemIf = false,
                                                                              StartTag = issueDate,
                                                                              Expression = "./IssueDate"
                                                                          },
                                                                  };
                return tableElements;
            };

            tableTag.MakeTableElementCallback = MakeTableElementCallback;

            var tableProcessor = new TableProcessor
            {
                TableTag = tableTag,
                DataReader = this.dataReader
            };
            tableProcessor.Process();

            var expectedTableStructure = new[]
                {
                new[] { "#", "Certificate", "Date" },
                new[] { string.Empty, string.Empty, "Issue", "Expiration" },
                new[] { "1", "2", "3", "4" },
                new[] { "1", "First static text", "01.04.2014", "Second static text" },
                new[] { "2", "First static text", "01.03.2014", "Second static text" },
                new[] { "3", "First static text", "01.01.2011", "Second static text" },
                new[] { "This", "row", "stays", "untouched" },
            };

            var rows = tableTag.Table.Elements(WordMl.TableRowName).ToList();
            Assert.AreEqual(expectedTableStructure.Count(), rows.Count());
            int rowIndex = 0;
            foreach (var row in rows)
            {
                var expectedRow = expectedTableStructure[rowIndex];

                var cellsInRow = row.Elements(WordMl.TableCellName).ToList();

                Assert.AreEqual(expectedRow.Count(), cellsInRow.Count());

                int cellIndex = 0;
                foreach (var cell in cellsInRow)
                {
                    Assert.AreEqual(expectedRow[cellIndex], cell.Value);
                    cellIndex++;
                }

                rowIndex++;
            }
            /*
            var tagsBetween =
                tableTag.TagTable.ElementsAfterSelf().Where(element => element.IsBefore(tableTag.TagContent));
            Assert.IsFalse(tagsBetween.Any());

            tagsBetween =
                tableTag.TagEndContent.ElementsAfterSelf().Where(element => element.IsBefore(tableTag.TagEndTable));
            Assert.IsFalse(tagsBetween.Any());*/
            this.ValidateTagsRemoved(root);
        }

        [TestMethod]
        public void TestProcessItemIfCells()
        {
            XElement root;
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "TableItemIfsTest.xml"))
            {
                var document = XDocument.Load(docStream);
                root = document.Root.Element(WordMl.BodyName);
            }
            var tableTag = this.GetTableTag(root);

            XDocument data;
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "TableItemIfsData.xml"))
            {
                data = XDocument.Load(docStream);
            }

            Func<XElement, IEnumerable<TableElement>> MakeTableElementCallback = element =>
            {
                var cells = element.Descendants(WordMl.TableCellName).ToList();
                ICollection<TableElement> tableElements = new Collection<TableElement>();
                foreach (var cell in cells)
                {
                    var cellTags = cell.Descendants(WordMl.SdtName).ToList();
                    var firtsIf = new TableElement
                                      {
                                          IsIndex = false,
                                          IsItem = false,
                                          IsItemIf = true,
                                          StartTag = cellTags[0],
                                          EndTag = cellTags[2],
                                          Expression = cellTags[0].Value,
                                          TagElements =
                                              new TableElement[]
                                                  {
                                                      new TableElement
                                                          {
                                                              IsIndex = false,
                                                              IsItemIf = false,
                                                              IsItem = true,
                                                              StartTag = cellTags[1],
                                                              Expression = cellTags[1].Value,
                                                          }
                                                  }
                                      };

                    var secondIf = new TableElement
                                       {
                                           IsIndex = false,
                                           IsItem = false,
                                           IsItemIf = true,
                                           StartTag = cellTags[3],
                                           EndTag = cellTags[5],
                                           Expression = cellTags[3].Value,
                                           TagElements =
                                               new TableElement[]
                                                   {
                                                       new TableElement
                                                           {
                                                               IsIndex = false,
                                                               IsItemIf = false,
                                                               IsItem = true,
                                                               StartTag = cellTags[4],
                                                               Expression = cellTags[4].Value,
                                                           }
                                                   }
                                       };

                    tableElements.Add(firtsIf);
                    tableElements.Add(secondIf);
                }

                return tableElements;
            };

            tableTag.MakeTableElementCallback = MakeTableElementCallback;

            var tableProcessor = new TableProcessor
            {
                TableTag = tableTag,
                DataReader = DataReaderFactory.CreateReader(data)
            };
            tableProcessor.Process();

            var expectedTableStructure = new[]
                {
                new[] { "#", "Certificate", "Date" },
                new[] { string.Empty, string.Empty, "Issue", "Expiration" },
                new[] { "1", "2", "3", "4" },
                new[] { "row1 - value1", "row1 - value2", "row1 - value3", "row1 - value4" },
                new[] { "row2 - value1", "row2 - value2", "row2 - value3", "row2 - value4" },
                new[] { "row3 - value1", "row3 - value2", "row3 - value3", "row3 - value4" },
                new[] { "row4 - value1", "row4 - value2", "row4 - value3", "row4 - value4" },
                new[] { "row5 - value1", "row5 - value2", "row5 - value3", "row5 - value4" },
                new[] { "This", "row", "stays", "untouched" },
            };

            var rows = tableTag.Table.Elements(WordMl.TableRowName).ToList();
            Assert.AreEqual(expectedTableStructure.Count(), rows.Count());
            int rowIndex = 0;
            foreach (var row in rows)
            {
                var expectedRow = expectedTableStructure[rowIndex];

                var cellsInRow = row.Elements(WordMl.TableCellName).ToList();

                Assert.AreEqual(expectedRow.Count(), cellsInRow.Count());

                int cellIndex = 0;
                foreach (var cell in cellsInRow)
                {
                    Assert.AreEqual(expectedRow[cellIndex], cell.Value);
                    cellIndex++;
                }

                rowIndex++;
            }
            /*
            var tagsBetween =
                tableTag.TagTable.ElementsAfterSelf().Where(element => element.IsBefore(tableTag.TagContent));
            Assert.IsFalse(tagsBetween.Any());

            tagsBetween =
                tableTag.TagEndContent.ElementsAfterSelf().Where(element => element.IsBefore(tableTag.TagEndTable));
            Assert.IsFalse(tagsBetween.Any());*/
            this.ValidateTagsRemoved(root);
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void TestProcessNullTag()
        {
            var processor = new TableProcessor();
            processor.DataReader = this.dataReader;
            processor.Process();
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void TestProcessNullReader()
        {
            var processor = new TableProcessor { TableTag = this.GetTableTag() };
            processor.Process();
        }

        private TableTag GetTableTag(XElement root)
        {
            XElement table = root.Element(WordMl.TableName);
            const string ItemsSource = "//Test/Certificates/Certificate";
            int? dynamicRowValue = 4;

            return new TableTag
                       {
                           Table = table,
                           ItemsSource = ItemsSource,
                           DynamicRow = dynamicRowValue,
                           TagTable = TraverseUtils.TagElement(root, "Table"),
                         //  TagContent = TraverseUtils.TagElement(root, "Content"),
                          // TagEndContent = TraverseUtils.TagElement(root, "EndContent"),
                           TagEndTable = TraverseUtils.TagElement(root, "EndTable"),
                       };
        }

        private TableTag GetTableTag()
        {
            return this.GetTableTag(this.documentRoot);
        }
    }
}
