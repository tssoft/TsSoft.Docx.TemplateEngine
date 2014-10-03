using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags.Processors
{
    [TestClass]
    public class TableProcessorTest
    {
        private XElement documentRoot;
        private DataReader dataReader;

        [TestInitialize]
        public void Initialize()
        {
            using (var templateStream = AssemblyResourceHelper.GetResourceStream(this, "TableProcessorTemplateTest.xml"))
            {

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(templateStream);
                using (XmlNodeReader nodeReader = new XmlNodeReader(xmlDoc))
                {
                    nodeReader.MoveToContent();
                    var document = XDocument.Load(nodeReader);
                    documentRoot = document.Root.Element(WordMl.BodyName);
                }
            }

            using (var dataStream = AssemblyResourceHelper.GetResourceStream(this, "TableProcessorDataTest.xml"))
            {
                var xmlDoc = new XmlDocument();
                xmlDoc.Load(dataStream);
                dataReader = DataReaderFactory.CreateReader(xmlDoc);
            }
        }

        [TestMethod]
        public void TestProcess()
        {
            var tableTag = GetTableTag();

            var tableProcessor = new TableProcessor
            {
                TableTag = tableTag,
                DataReader = dataReader
            };
            tableProcessor.Process();

            var expectedTableStructure = new string[][]
            {
                new string[] { "#", "Certificate", "Date" },
                new string[] { string.Empty, string.Empty, "Issue", "Expiration" },
                new string[] { "1", "2", "3", "4" },
                new string[] { "1", "Информатика", "01.04.2014", "01.10.2015" },
                new string[] { "2", "Математика", "01.03.2014", "01.09.2015" },
                new string[] { "3", "Языки программирования", "01.01.2011", "01.01.2012" },
                new string[] { "This", "row", "stays", "untouched" },
            };

            var rows = tableTag.Table.Elements(WordMl.TableRowName);
            Assert.AreEqual(expectedTableStructure.Count(), rows.Count());
            int rowIndex = 0;
            foreach (var row in rows)
            {
                var expectedRow = expectedTableStructure[rowIndex];

                var cellsInRow = row.Elements(WordMl.TableCellName);

                Assert.AreEqual(expectedRow.Count(), cellsInRow.Count());

                int cellIndex = 0;
                foreach (var cell in cellsInRow)
                {
                    Assert.AreEqual(expectedRow[cellIndex], cell.Value);
                    cellIndex++;
                }

                rowIndex++;
            }

            Assert.IsNull(tableTag.TagTable.Parent);
            Assert.IsNull(tableTag.TagContent.Parent);
            var tagsBetween =
                tableTag.TagTable.ElementsAfterSelf().Where(element => element.IsBefore(tableTag.TagContent));
            Assert.IsFalse(tagsBetween.Any());
            Assert.IsNull(tableTag.TagEndContent.Parent);
            Assert.IsNull(tableTag.TagEndTable.Parent);
            tagsBetween =
                tableTag.TagEndContent.ElementsAfterSelf().Where(element => element.IsBefore(tableTag.TagEndTable));
            Assert.IsFalse(tagsBetween.Any());
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void TestProcessNullTag()
        {
            var processor = new TableProcessor();
            processor.DataReader = dataReader;
            processor.Process();
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void TestProcessNullReader()
        {
            var processor = new TableProcessor();
            processor.TableTag = GetTableTag();
            processor.Process();
        }

        private TableTag GetTableTag()
        {
            XElement table = documentRoot.Element(WordMl.TableName);
            var itemsSource = "//Test/Certificates/Certificate";
            int? dynamicRowValue = 4;

            return new TableTag()
            {
                Table = table,
                ItemsSource = itemsSource,
                DynamicRow = dynamicRowValue,
                TagTable = TagElement(documentRoot, "Table"),
                TagContent = TagElement(documentRoot, "Content"),
                TagEndContent = TagElement(documentRoot, "EndContent"),
                TagEndTable = TagElement(documentRoot, "EndTable"),
            };
        }

        private XElement TagElement(XElement root, string nameTag)
        {
            return
                root.Elements(WordMl.SdtName)
                    .First(
                        element =>
                        element.Element(WordMl.SdtPrName)
                               .Element(WordMl.TagName)
                               .Attribute(WordMl.ValAttributeName)
                               .Value == nameTag);
        }
    }
}
