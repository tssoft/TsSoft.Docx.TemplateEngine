using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Parsers;
using TsSoft.Docx.TemplateEngine.Tags;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    /// <summary>
    /// Test TableParser class
    /// </summary>
    /// <author>Георгий Поликарпов</author>
    [TestClass]
    public class TableParserTest
    {
        private XElement documentRoot;

        [TestInitialize]
        public void Initialize()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "table.xml");
            var doc = XDocument.Load(docStream);
            documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");
        }

        [TestMethod]
        public void DoVariousOrderTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var pElement = root.Element(WordMl.WordMlNamespace + "p");
            var itemsElement = TraverseUtils.TagElement(root, "Items");
            var itemsSource = itemsElement.Value;
            itemsElement.AddBeforeSelf(pElement);
            var dynamicRowElement = TraverseUtils.TagElement(root, "DynamicRow");
            var dynamicRowValue = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            dynamicRowElement.AddBeforeSelf(pElement);
            var contentElement = TraverseUtils.TagElement(root, "Content");
            contentElement.AddBeforeSelf(pElement);
            var endContentElement = TraverseUtils.TagElement(root, "EndContent");
            endContentElement.AddBeforeSelf(pElement);
            var tableElement = root.Element(WordMl.WordMlNamespace + "tbl");
            tableElement.AddBeforeSelf(pElement);
            var endTableElement = TraverseUtils.TagElement(root, "EndTable");
            endTableElement.AddBeforeSelf(pElement);
            var startElement = TraverseUtils.TagElement(root, "Table");

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
            var dynamicRowElement = TraverseUtils.TagElement(root, "DynamicRow");
            int dynamicRowValue = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            var itemsSource = TraverseUtils.TagElement(root, "Items").Value;
            var tableElement = root.Element(WordMl.WordMlNamespace + "tbl");
            var newTableElement = new XElement(tableElement);
            newTableElement.Elements(WordMl.WordMlNamespace + "tr").Last().Remove();
            tableElement.AddAfterSelf(newTableElement);

            var startElement = TraverseUtils.TagElement(root, "Table");

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
            var dynamicRowElement = TraverseUtils.TagElement(root, "DynamicRow");
            int dynamicRowValue = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            var newDynamicRowElement = new XElement(dynamicRowElement);
            SetTagElementValue(newDynamicRowElement, (dynamicRowValue + 10).ToString());
            dynamicRowElement.AddAfterSelf(newDynamicRowElement);
            var itemsSource = TraverseUtils.TagElement(root, "Items").Value;
            var tableElement = root.Element(WordMl.WordMlNamespace + "tbl");
            var startElement = TraverseUtils.TagElement(root, "Table");

            var result = parser.Do(startElement);
            Assert.AreEqual(dynamicRowValue, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.AreEqual(tableElement, result.Table);

            root = new XElement(documentRoot);
            dynamicRowElement = TraverseUtils.TagElement(root, "DynamicRow");
            dynamicRowValue = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            var itemsElement = TraverseUtils.TagElement(root, "Items");
            itemsSource = itemsElement.Value;
            var newItemsElement = new XElement(itemsElement);
            SetTagElementValue(newItemsElement, @"//wrongPath");
            itemsElement.AddAfterSelf(newItemsElement);
            tableElement = root.Element(WordMl.WordMlNamespace + "tbl");
            startElement = TraverseUtils.TagElement(root, "Table");

            result = parser.Do(startElement);
            Assert.AreEqual(dynamicRowValue, result.DynamicRow);
            Assert.AreEqual(itemsSource, result.ItemsSource);
            Assert.AreEqual(tableElement, result.Table);

            root = new XElement(documentRoot);
            dynamicRowElement = TraverseUtils.TagElement(root, "DynamicRow");
            dynamicRowValue = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            itemsSource = TraverseUtils.TagElement(root, "Items").Value;
            var contentElement = TraverseUtils.TagElement(root, "Content");
            var newContentElement = new XElement(contentElement);
            tableElement = root.Element(WordMl.WordMlNamespace + "tbl");
            var newTableElement = new XElement(tableElement);
            newTableElement.Elements(WordMl.WordMlNamespace + "tr").Last().Remove();
            var endContentElement = TraverseUtils.TagElement(root, "EndContent");
            var newEndContentElement = new XElement(endContentElement);
            endContentElement.AddBeforeSelf(newEndContentElement);
            endContentElement.AddBeforeSelf(newContentElement);
            endContentElement.AddBeforeSelf(newTableElement);

            startElement = TraverseUtils.TagElement(root, "Table");

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
            var dynamicRowElement = TraverseUtils.TagElement(root, "DynamicRow");
            var dynamicRowValue = (dynamicRowElement.Value == "") ? 0 : int.Parse(dynamicRowElement.Value);
            var itemsSource = TraverseUtils.TagElement(root, "Items").Value;
            var tableElement = root.Element(WordMl.WordMlNamespace + "tbl");
            tableElement.Remove();
            TraverseUtils.TagElement(root, "DynamicRow").AddAfterSelf(tableElement);

            var startElement = TraverseUtils.TagElement(root, "Table");

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
            SetTagElementValue(TraverseUtils.TagElement(root, "DynamicRow"), "");
            var itemsSource = TraverseUtils.TagElement(root, "Items").Value;
            var tableElement = root.Element(WordMl.WordMlNamespace + "tbl");

            var startElement = TraverseUtils.TagElement(root, "Table");

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
            var itemsElement = TraverseUtils.TagElement(root, "Items");
            SetTagElementValue(itemsElement, "");
            var startElement = TraverseUtils.TagElement(root, "Table");

            parser.Do(startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void DoMissingContentTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TraverseUtils.TagElement(root, "Content").Remove();
            var startElement = TraverseUtils.TagElement(root, "Table");
            parser.Do(startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void DoMissingEndContentTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TraverseUtils.TagElement(root, "EndContent").Remove();
            var startElement = TraverseUtils.TagElement(root, "Table");
            parser.Do(startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void DoMissingEndTableTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TraverseUtils.TagElement(root, "EndTable").Remove();
            var startElement = TraverseUtils.TagElement(root, "Table");
            parser.Do(startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void DoMissingItemsTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TraverseUtils.TagElement(root, "Items").Remove();
            var startElement = TraverseUtils.TagElement(root, "Table");
            parser.Do(startElement);
        }

        [TestMethod]
        public void DoMissingDynamicRowTagTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TraverseUtils.TagElement(root, "DynamicRow").Remove();
            var itemsSource = TraverseUtils.TagElement(root, "Items").Value;
            var tableElement = root.Element(WordMl.WordMlNamespace + "tbl");

            var startElement = TraverseUtils.TagElement(root, "Table");

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
            var contentElement = TraverseUtils.TagElement(root, "Content");
            var newContentElement = new XElement(contentElement);
            TraverseUtils.TagElement(root, "EndContent").AddBeforeSelf(newContentElement);
            var startElement = TraverseUtils.TagElement(root, "Table");
            parser.Do(startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void DoTwoTableBeforeEndTableTest()
        {
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var startElement = TraverseUtils.TagElement(root, "Table");
            var newTableTagElement = new XElement(startElement);
            TraverseUtils.TagElement(root, "EndTable").AddBeforeSelf(newTableTagElement);
            parser.Do(startElement);
        }



        private void SetTagElementValue(XElement element, string value)
        {
            element.Element(WordMl.SdtContentName).Element(WordMl.WordMlNamespace + "p").Element(WordMl.WordMlNamespace + "r").Element(WordMl.WordMlNamespace + "t").Value = value;
        }
    }
}
