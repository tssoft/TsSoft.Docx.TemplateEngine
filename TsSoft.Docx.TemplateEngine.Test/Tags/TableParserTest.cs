using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    /// <summary>
    /// Test TableParser class
    /// </summary>
    [TestClass]
    public class TableParserTest
    {
        private XElement documentRoot;

        private XElement nestedDocumentRoot;
        
        [TestInitialize]
        public void Initialize()
        {
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "TableParserTest.xml"))
            {
                var doc = XDocument.Load(docStream);
                documentRoot = doc.Root.Element(WordMl.BodyName);
            }
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "TableParserNestedTest.xml"))
            {
                var doc = XDocument.Load(docStream);
                nestedDocumentRoot = doc.Root.Element(WordMl.BodyName);
            }
        }

        [TestMethod]
        public void TestParseVariousOrderTag()
        {
            var processorMock = new TagProcessorMock<TableProcessor>();
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var pElement = root.Element(WordMl.ParagraphName);
            var itemsElement = TraverseUtils.TagElement(root, "Items");
            itemsElement.AddBeforeSelf(pElement);
            var dynamicRowElement = TraverseUtils.TagElement(root, "DynamicRow");
            dynamicRowElement.AddBeforeSelf(pElement);
            var contentElement = TraverseUtils.TagElement(root, "Content");
            contentElement.AddBeforeSelf(pElement);
            var endContentElement = TraverseUtils.TagElement(root, "EndContent");
            endContentElement.AddBeforeSelf(pElement);
            var tableElement = root.Element(WordMl.TableName);
            tableElement.AddBeforeSelf(pElement);
            var endTableElement = TraverseUtils.TagElement(root, "EndTable");
            endTableElement.AddBeforeSelf(pElement);
            var startElement = TraverseUtils.TagElement(root, "Table");

            parser.Parse(processorMock, startElement);
            var processor = processorMock.InnerProcessor;
            var tag = processor.TableTag;
            Assert.AreEqual(4, tag.DynamicRow);
            Assert.AreEqual("//test/certificates", tag.ItemsSource);
            Assert.AreEqual(tableElement, tag.Table);
            CheckTagElements(tag);

            itemsElement.Remove();
            endTableElement.AddBeforeSelf(itemsElement);

            parser.Parse(processorMock, startElement);
            processor = processorMock.InnerProcessor;
            tag = processor.TableTag;
            Assert.AreEqual(4, tag.DynamicRow);
            Assert.AreEqual("//test/certificates", tag.ItemsSource);
            Assert.AreEqual(tableElement, tag.Table);
            CheckTagElements(tag);

            dynamicRowElement.Remove();
            endTableElement.AddBeforeSelf(dynamicRowElement);

            parser.Parse(processorMock, startElement);
            processor = processorMock.InnerProcessor;
            tag = processor.TableTag;
            Assert.AreEqual(4, tag.DynamicRow);
            Assert.AreEqual("//test/certificates", tag.ItemsSource);
            Assert.AreEqual(tableElement, tag.Table);
            CheckTagElements(tag);

            contentElement.Remove();
            tableElement.Remove();
            endContentElement.Remove();
            endTableElement.AddBeforeSelf(contentElement);
            endTableElement.AddBeforeSelf(tableElement);
            endTableElement.AddBeforeSelf(endContentElement);

            parser.Parse(processorMock, startElement);
            processor = processorMock.InnerProcessor;
            tag = processor.TableTag;
            Assert.AreEqual(4, tag.DynamicRow);
            Assert.AreEqual("//test/certificates", tag.ItemsSource);
            Assert.AreEqual(tableElement, tag.Table);
            CheckTagElements(tag);
        }

        [TestMethod]
        public void TestParseDoubleTable()
        {
            var processorMock = new TagProcessorMock<TableProcessor>();
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var tableElement = root.Element(WordMl.TableName);
            var newTableElement = new XElement(tableElement);
            newTableElement.Elements(WordMl.TableRowName).Last().Remove();
            tableElement.AddAfterSelf(newTableElement);

            var startElement = TraverseUtils.TagElement(root, "Table");

            parser.Parse(processorMock, startElement);
            var processor = processorMock.InnerProcessor;
            var tag = processor.TableTag;
            Assert.AreEqual(4, tag.DynamicRow);
            Assert.AreEqual("//test/certificates", tag.ItemsSource);
            Assert.AreEqual(tableElement, tag.Table);
            CheckTagElements(tag);
        }

        [TestMethod]
        public void TestParseDoubleTag()
        {
            var processorMock = new TagProcessorMock<TableProcessor>();
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var dynamicRowElement = TraverseUtils.TagElement(root, "DynamicRow");
            var newDynamicRowElement = new XElement(dynamicRowElement);
            SetTagElementValue(newDynamicRowElement, "10");
            dynamicRowElement.AddAfterSelf(newDynamicRowElement);
            var tableElement = root.Element(WordMl.TableName);
            var startElement = TraverseUtils.TagElement(root, "Table");

            parser.Parse(processorMock, startElement);
            var processor = processorMock.InnerProcessor;
            var tag = processor.TableTag;
            Assert.AreEqual(4, tag.DynamicRow);
            Assert.AreEqual("//test/certificates", tag.ItemsSource);
            Assert.AreEqual(tableElement, tag.Table);
            CheckTagElements(tag);

            root = new XElement(documentRoot);
            var itemsElement = TraverseUtils.TagElement(root, "Items");
            var newItemsElement = new XElement(itemsElement);
            SetTagElementValue(newItemsElement, @"//wrongPath");
            itemsElement.AddAfterSelf(newItemsElement);
            tableElement = root.Element(WordMl.TableName);
            startElement = TraverseUtils.TagElement(root, "Table");

            parser.Parse(processorMock, startElement);
            processor = processorMock.InnerProcessor;
            tag = processor.TableTag;
            Assert.AreEqual(4, tag.DynamicRow);
            Assert.AreEqual("//test/certificates", tag.ItemsSource);
            Assert.AreEqual(tableElement, tag.Table);
            CheckTagElements(tag);

            root = new XElement(documentRoot);
            var contentElement = TraverseUtils.TagElement(root, "Content");
            var newContentElement = new XElement(contentElement);
            tableElement = root.Element(WordMl.TableName);
            var newTableElement = new XElement(tableElement);
            newTableElement.Elements(WordMl.TableRowName).Last().Remove();
            var endContentElement = TraverseUtils.TagElement(root, "EndContent");
            var newEndContentElement = new XElement(endContentElement);
            endContentElement.AddBeforeSelf(newEndContentElement);
            endContentElement.AddBeforeSelf(newContentElement);
            endContentElement.AddBeforeSelf(newTableElement);

            startElement = TraverseUtils.TagElement(root, "Table");

            parser.Parse(processorMock, startElement);
            processor = processorMock.InnerProcessor;
            tag = processor.TableTag;
            Assert.AreEqual(4, tag.DynamicRow);
            Assert.AreEqual("//test/certificates", tag.ItemsSource);
            Assert.AreEqual(tableElement, tag.Table);
            CheckTagElements(tag);
        }

        [TestMethod]
        public void TestParseEmptyContent()
        {
            var processorMock = new TagProcessorMock<TableProcessor>();
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var tableElement = root.Element(WordMl.TableName);
            tableElement.Remove();
            TraverseUtils.TagElement(root, "DynamicRow").AddAfterSelf(tableElement);

            var startElement = TraverseUtils.TagElement(root, "Table");

            parser.Parse(processorMock, startElement);
            var processor = processorMock.InnerProcessor;
            var tag = processor.TableTag;
            Assert.AreEqual(4, tag.DynamicRow);
            Assert.AreEqual("//test/certificates", tag.ItemsSource);
            Assert.IsNull(tag.Table);
            CheckTagElements(tag);
        }

        [TestMethod]
        public void TestParseEmptyDynamicRow()
        {
            var processorMock = new TagProcessorMock<TableProcessor>();
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var dynamicRowElement = TraverseUtils.TagElement(root, "DynamicRow");
            SetTagElementValue(dynamicRowElement, string.Empty);
            var tableElement = root.Element(WordMl.TableName);

            var startElement = TraverseUtils.TagElement(root, "Table");

            parser.Parse(processorMock, startElement);
            var processor = processorMock.InnerProcessor;
            var tag = processor.TableTag;
            Assert.IsFalse(tag.DynamicRow.HasValue);
            Assert.AreEqual("//test/certificates", tag.ItemsSource);
            Assert.AreEqual(tableElement, tag.Table);
            CheckTagElements(tag);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void TestParseEmptyItems()
        {
            var processorMock = new TagProcessorMock<TableProcessor>();
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var itemsElement = TraverseUtils.TagElement(root, "Items");
            SetTagElementValue(itemsElement, string.Empty);
            var startElement = TraverseUtils.TagElement(root, "Table");
            parser.Parse(processorMock, startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void TestParseMissingContentTag()
        {
            var processorMock = new TagProcessorMock<TableProcessor>();
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TraverseUtils.TagElement(root, "Content").Remove();
            var startElement = TraverseUtils.TagElement(root, "Table");
            parser.Parse(processorMock, startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void TestParseMissingEndContentTag()
        {
            var processorMock = new TagProcessorMock<TableProcessor>();
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TraverseUtils.TagElement(root, "EndContent").Remove();
            var startElement = TraverseUtils.TagElement(root, "Table");
            parser.Parse(processorMock, startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void TestParseMissingEndTableTag()
        {
            var processorMock = new TagProcessorMock<TableProcessor>();
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TraverseUtils.TagElement(root, "EndTable").Remove();
            var startElement = TraverseUtils.TagElement(root, "Table");
            parser.Parse(processorMock, startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void TestParseMissingItemsTag()
        {
            var processorMock = new TagProcessorMock<TableProcessor>();
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TraverseUtils.TagElement(root, "Items").Remove();
            var startElement = TraverseUtils.TagElement(root, "Table");
            parser.Parse(processorMock, startElement);
        }

        [TestMethod]
        public void TestParseMissingDynamicRowTag()
        {
            var processorMock = new TagProcessorMock<TableProcessor>();
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            TraverseUtils.TagElement(root, "DynamicRow").Remove();
            var tableElement = root.Element(WordMl.TableName);

            var startElement = TraverseUtils.TagElement(root, "Table");

            parser.Parse(processorMock, startElement);
            var processor = processorMock.InnerProcessor;
            var tag = processor.TableTag;
            Assert.IsFalse(tag.DynamicRow.HasValue);
            Assert.AreEqual("//test/certificates", tag.ItemsSource);
            Assert.AreEqual(tableElement, tag.Table);
            CheckTagElements(tag);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void TestParseNullProcess()
        {
            var root = new XElement(documentRoot);
            var startElement = TraverseUtils.TagElement(root, "Table");
            var parser = new TableParser();
            parser.Parse(null, startElement);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void TestParseTwoTableBeforeEndTable()
        {
            var processorMock = new TagProcessorMock<TableProcessor>();
            var parser = new TableParser();

            var root = new XElement(documentRoot);
            var startElement = TraverseUtils.TagElement(root, "Table");
            var newTableTagElement = new XElement(startElement);
            TraverseUtils.TagElement(root, "EndTable").AddBeforeSelf(newTableTagElement);
            parser.Parse(processorMock, startElement);
        }

        [TestMethod]
        public void TestParseNedted()
        {
            var parser = new TableParser();
            RootProcessor rootProcessor = new RootProcessor();
            var tableElement = nestedDocumentRoot.Element(WordMl.TableName);
            parser.Parse(rootProcessor, nestedDocumentRoot.Elements(WordMl.SdtName).First());

            var tableProcessor = rootProcessor.Processors.First();
            var tag = ((TableProcessor)tableProcessor).TableTag;
            Assert.AreEqual(4, tag.DynamicRow);
            Assert.AreEqual("//test/certificates", tag.ItemsSource);
            Assert.AreEqual(tableElement, tag.Table);
            CheckTagElements(tag);
            Assert.AreEqual(2, tableProcessor.Processors.Count);
            Assert.IsTrue(tableProcessor.Processors.All(p => p is TextProcessor));
        }

        private void SetTagElementValue(XElement element, string value)
        {
            element.Element(WordMl.SdtContentName).Element(WordMl.ParagraphName).Element(WordMl.WordMlNamespace + "r").Element(WordMl.WordMlNamespace + "t").Value = value;
        }

        private void CheckTagElements(TableTag tag)
        {
            Assert.IsNotNull(tag.TagTable);
            Assert.AreEqual("Table", tag.TagTable.Element(WordMl.SdtPrName).Element(WordMl.TagName).Attribute(WordMl.ValAttributeName).Value);
            Assert.IsNotNull(tag.TagContent);
            Assert.AreEqual("Content", tag.TagContent.Element(WordMl.SdtPrName).Element(WordMl.TagName).Attribute(WordMl.ValAttributeName).Value);
            Assert.IsNotNull(tag.TagEndContent);
            Assert.AreEqual("EndContent", tag.TagEndContent.Element(WordMl.SdtPrName).Element(WordMl.TagName).Attribute(WordMl.ValAttributeName).Value);
            Assert.IsNotNull(tag.TagEndTable);
            Assert.AreEqual("EndTable", tag.TagEndTable.Element(WordMl.SdtPrName).Element(WordMl.TagName).Attribute(WordMl.ValAttributeName).Value);

        }
    }
}
