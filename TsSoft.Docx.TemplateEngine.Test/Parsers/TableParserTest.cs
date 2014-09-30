using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Parsers;

namespace TsSoft.Docx.TemplateEngine.Test.Parsers
{
    /// <summary>
    /// Test TableParser class
    /// </summary>
    /// <author>Георгий Поликарпов</author>
    [TestClass]
    public class TableParserTest
    {
        public static readonly XNamespace WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XName SdtName = WordMlNamespace + "sdt";
        public static readonly XName SdtPrName = WordMlNamespace + "sdtPr";
        public static readonly XName TagName = WordMlNamespace + "tag";
        public static readonly XName SdtContentName = WordMlNamespace + "sdtContent";
        public static readonly XName ValAttributeName = WordMlNamespace + "val";
        private XElement documentRoot;

        [TestInitialize]
        public void Initialize()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load("../../Parsers/table.xml");
            using (XmlNodeReader nodeReader = new XmlNodeReader(xmlDoc))
            {
                nodeReader.MoveToContent();
                var document = XDocument.Load(nodeReader);

                documentRoot = document.Root.Element(WordMlNamespace + "body");
            }
        }

        [TestMethod]
        public void DoVariousOrderTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var pElement = root.Element(WordMlNamespace + "p");
            var itemsElement = TagElement(root, "Items");
            var itemsSource = itemsElement.Value;
            itemsElement.AddBeforeSelf(pElement);
            var dynamicRowElement = TagElement(root, "DynamicRow");
            var dynamicRowValue = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            dynamicRowElement.AddBeforeSelf(pElement);
            var contentElement = TagElement(root, "Content");
            contentElement.AddBeforeSelf(pElement);
            var endContentElement = TagElement(root, "EndContent");
            endContentElement.AddBeforeSelf(pElement);
            var tableElement = root.Element(WordMlNamespace + "tbl");
            tableElement.AddBeforeSelf(pElement);
            var endTableElement = TagElement(root, "EndTable");
            endTableElement.AddBeforeSelf(pElement);
            var startElement = TagElement(root, "Table");

            var result = parser.Do(startElement);
            Assert.AreEqual(dynamicRowValue, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.AreEqual(tableElement, result.Table);

            itemsElement.Remove();
            endTableElement.AddBeforeSelf(itemsElement);

            result = parser.Do(startElement);
            Assert.AreEqual(dynamicRowValue, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.AreEqual(tableElement, result.Table);

            dynamicRowElement.Remove();
            endTableElement.AddBeforeSelf(dynamicRowElement);

            result = parser.Do(startElement);
            Assert.AreEqual(dynamicRowValue, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.AreEqual(tableElement, result.Table);

            contentElement.Remove();
            tableElement.Remove();
            endContentElement.Remove();
            endTableElement.AddBeforeSelf(contentElement);
            endTableElement.AddBeforeSelf(tableElement);
            endTableElement.AddBeforeSelf(endContentElement);

            result = parser.Do(startElement);
            Assert.AreEqual(dynamicRowValue, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.AreEqual(tableElement, result.Table);
        }

        [TestMethod]
        public void DoDoubleTableTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var dynamicRowElement = TagElement(root, "DynamicRow");
            int dynamicRowValue = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            var itemsSource = TagElement(root, "Items").Value;
            var tableElement = root.Element(WordMlNamespace + "tbl");
            var newTableElement = new XElement(tableElement);
            newTableElement.Elements(WordMlNamespace + "tr").Last().Remove();
            tableElement.AddAfterSelf(newTableElement);

            var startElement = TagElement(root, "Table");

            var result = parser.Do(startElement);
            Assert.AreEqual(dynamicRowValue, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.AreEqual(tableElement, result.Table);
        }

        [TestMethod]
        public void DoDoubleTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var dynamicRowElement = TagElement(root, "DynamicRow");
            int dynamicRowValue = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            var newDynamicRowElement = new XElement(dynamicRowElement);
            SetTagElementValue(newDynamicRowElement, (dynamicRowValue + 10).ToString());
            dynamicRowElement.AddAfterSelf(newDynamicRowElement);
            var itemsSource = TagElement(root, "Items").Value;
            var tableElement = root.Element(WordMlNamespace + "tbl");
            var startElement = TagElement(root, "Table");

            var result = parser.Do(startElement);
            Assert.AreEqual(dynamicRowValue, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.AreEqual(tableElement, result.Table);

            root = new XElement(documentRoot);
            dynamicRowElement = TagElement(root, "DynamicRow");
            dynamicRowValue = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            var itemsElement = TagElement(root, "Items");
            itemsSource = itemsElement.Value;
            var newItemsElement = new XElement(itemsElement);
            SetTagElementValue(newItemsElement, @"//wrongPath");
            itemsElement.AddAfterSelf(newItemsElement);
            tableElement = root.Element(WordMlNamespace + "tbl");
            startElement = TagElement(root, "Table");

            result = parser.Do(startElement);
            Assert.AreEqual(dynamicRowValue, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.AreEqual(tableElement, result.Table);

            root = new XElement(documentRoot);
            dynamicRowElement = TagElement(root, "DynamicRow");
            dynamicRowValue = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            itemsSource = TagElement(root, "Items").Value;
            var contentElement = TagElement(root, "Content");
            var newContentElement = new XElement(contentElement);
            tableElement = root.Element(WordMlNamespace + "tbl");
            var newTableElement = new XElement(tableElement);
            newTableElement.Elements(WordMlNamespace + "tr").Last().Remove();
            var endContentElement = TagElement(root, "EndContent");
            var newEndContentElement = new XElement(endContentElement);
            endContentElement.AddBeforeSelf(newEndContentElement);
            endContentElement.AddBeforeSelf(newContentElement);
            endContentElement.AddBeforeSelf(newTableElement);

            startElement = TagElement(root, "Table");

            result = parser.Do(startElement);
            Assert.AreEqual(dynamicRowValue, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.AreEqual(tableElement, result.Table);
        }

        [TestMethod]
        public void DoEmptyContentTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var dynamicRowElement = TagElement(root, "DynamicRow");
            var dynamicRowValue = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            var itemsSource = TagElement(root, "Items").Value;
            var tableElement = root.Element(WordMlNamespace + "tbl");
            tableElement.Remove();
            TagElement(root, "DynamicRow").AddAfterSelf(tableElement);

            var startElement = TagElement(root, "Table");

            var result = parser.Do(startElement);
            Assert.AreEqual(dynamicRowValue, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.IsNull(result.Table);
        }

        [TestMethod]
        public void DoEmptyDynamicRowTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            SetTagElementValue(TagElement(root, "DynamicRow"), "");
            var itemsSource = TagElement(root, "Items").Value;
            var tableElement = root.Element(WordMlNamespace + "tbl");

            var startElement = TagElement(root, "Table");

            var result = parser.Do(startElement);
            Assert.AreEqual(0, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.AreEqual(tableElement, result.Table);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void DoEmptyItemsTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var itemsElement = TagElement(root, "Items");
            SetTagElementValue(itemsElement, "");
            var startElement = TagElement(root, "Table");

            parser.Do(startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void DoMissingContentTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TagElement(root, "Content").Remove();
            var startElement = TagElement(root, "Table");
            parser.Do(startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void DoMissingEndContentTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TagElement(root, "EndContent").Remove();
            var startElement = TagElement(root, "Table");
            parser.Do(startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void DoMissingEndTableTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TagElement(root, "EndTable").Remove();
            var startElement = TagElement(root, "Table");
            parser.Do(startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void DoMissingItemsTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TagElement(root, "Items").Remove();
            var startElement = TagElement(root, "Table");
            parser.Do(startElement);
        }

        [TestMethod]
        public void DoMissingDynamicRowTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TagElement(root, "DynamicRow").Remove();
            var itemsSource = TagElement(root, "Items").Value;
            var tableElement = root.Element(WordMlNamespace + "tbl");

            var startElement = TagElement(root, "Table");

            var result = parser.Do(startElement);
            Assert.AreEqual(0, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.AreEqual(tableElement, result.Table);

        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void DoNullArgumentTest()
        {
            var parser = new TableParser();
            parser.Do(null);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void DoTwoContentBeforeEndContentTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var contentElement = TagElement(root, "Content");
            var newContentElement = new XElement(contentElement);
            TagElement(root, "EndContent").AddBeforeSelf(newContentElement);
            var startElement = TagElement(root, "Table");
            parser.Do(startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void DoTwoTableBeforeEndTableTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var startElement = TagElement(root, "Table");
            var newTableTagElement = new XElement(startElement);
            TagElement(root, "EndTable").AddBeforeSelf(newTableTagElement);
            parser.Do(startElement);
        }

        private XElement TagElement(XElement root, string tagName)
        {
            return root.Elements(SdtName).Where(element => element.Element(SdtPrName).Element(TagName).Attribute(ValAttributeName).Value == tagName).First();
        }

        private void SetTagElementValue(XElement element, string value)
        {
            element.Element(SdtContentName).Element(WordMlNamespace + "p").Element(WordMlNamespace + "r").Element(WordMlNamespace + "t").Value = value;
        }
    }
}
