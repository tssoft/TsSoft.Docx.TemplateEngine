using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags.Processors
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;

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
                             .Equals("item") && e.Element(WordMl.SdtContentName).Value == "./Subject");
                    var issueDate =
                        tags.Find(
                            e =>
                            e.Element(WordMl.SdtPrName)
                             .Element(WordMl.TagName)
                             .Attribute(WordMl.ValAttributeName)
                             .Value.ToLower()
                             .Equals("item") && e.Element(WordMl.SdtContentName).Value == "./IssueDate");
                    var expireDate =
                        tags.Find(
                            e =>
                            e.Element(WordMl.SdtPrName)
                             .Element(WordMl.TagName)
                             .Attribute(WordMl.ValAttributeName)
                             .Value.ToLower()
                             .Equals("item") && e.Element(WordMl.SdtContentName).Value == "./ExpireDate");
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

            var expectedTableStructure = new[]
                {
                new[] { "#", "Certificate", "Date" },
                new[] { string.Empty, string.Empty, "Issue", "Expiration" },
                new[] { "1", "2", "3", "4" },
                new[] { "1", "Информатика", "01.04.2014", "01.10.2015" },
                new[] { "2", "Математика", "01.03.2014", "01.09.2015" },
                new[] { "3", "Языки программирования", "01.01.2011", "01.01.2012" },
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

            var tagsBetween =
                tableTag.TagTable.ElementsAfterSelf().Where(element => element.IsBefore(tableTag.TagContent));
            Assert.IsFalse(tagsBetween.Any());

            tagsBetween =
                tableTag.TagEndContent.ElementsAfterSelf().Where(element => element.IsBefore(tableTag.TagEndTable));
            Assert.IsFalse(tagsBetween.Any());
            this.ValidateTagsRemoved(this.documentRoot);
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

        private TableTag GetTableTag()
        {
            XElement table = this.documentRoot.Element(WordMl.TableName);
            const string ItemsSource = "//Test/Certificates/Certificate";
            int? dynamicRowValue = 4;

            return new TableTag
                       {
                           Table = table,
                           ItemsSource = ItemsSource,
                           DynamicRow = dynamicRowValue,
                           TagTable = TraverseUtils.TagElement(this.documentRoot, "Table"),
                           TagContent = TraverseUtils.TagElement(this.documentRoot, "Content"),
                           TagEndContent = TraverseUtils.TagElement(this.documentRoot, "EndContent"),
                           TagEndTable = TraverseUtils.TagElement(this.documentRoot, "EndTable"),
                       };
        }
    }
}
