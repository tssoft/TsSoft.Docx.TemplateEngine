using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    [TestClass]
    public class RepeaterParserTest
    {
        private XElement documentRoot;

        private XElement nestedDocumentRoot;

        [TestInitialize]
        public void Initialize()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "Repeater_Ok.xml");
            var doc = XDocument.Load(docStream);
            this.documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

            docStream = AssemblyResourceHelper.GetResourceStream(this, "Repeater_Nested.xml");
            doc = XDocument.Load(docStream);
            this.nestedDocumentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");
        }

        [TestMethod]
        public void TestRepeaterNotClosed()
        {
            const string TagName = "EndRepeater";
            var endRepeater = TraverseUtils.TagElement(this.documentRoot, TagName);
            endRepeater.Remove();
            var parser = new RepeaterParser();
            try
            {
                parser.Parse(new TagProcessorMock<RepeaterProcessor>(), this.documentRoot.Descendants(WordMl.SdtName).First());
                Assert.Fail("An exception shoud've been thrown");
            }
            catch (Exception e)
            {
                Assert.AreEqual(string.Format(MessageStrings.TagNotFoundOrEmpty, TagName), e.Message);
            }
        }
        /* НЕ АКТУАЛЕН
        [TestMethod]
        public void TestContentNotClosed()
        {
            const string TagName = "EndContent";
            var endContent = TraverseUtils.TagElement(this.documentRoot, TagName);
            endContent.Remove();
            var parser = new RepeaterParser();
            try
            {
                parser.Parse(new TagProcessorMock<RepeaterProcessor>(), this.documentRoot.Descendants(WordMl.SdtName).First());
                Assert.Fail("An exception shoud've been thrown");
            }
            catch (Exception e)
            {
                Assert.AreEqual(string.Format(MessageStrings.TagNotFoundOrEmpty, TagName), e.Message);
            }
        }*/
        /*
        [TestMethod]
        public void TestNoItems()
        {
            const string TagName = "Items";
            var items = TraverseUtils.TagElement(this.documentRoot, TagName);
            items.Remove();
            var parser = new RepeaterParser();
            try
            {
                parser.Parse(new TagProcessorMock<RepeaterProcessor>(), this.documentRoot.Descendants(WordMl.SdtName).First());

                Assert.Fail("An exception shoud've been thrown");
            }
            catch (Exception e)
            {
                Assert.AreEqual(string.Format(MessageStrings.TagNotFoundOrEmpty, TagName), e.Message);
            }
        }
         */

        [TestMethod]
        public void TestOkay()
        {
            var parser = new RepeaterParser();
            var tagProcessorMock = new TagProcessorMock<RepeaterProcessor>();
            parser.Parse(tagProcessorMock, this.documentRoot.Descendants(WordMl.SdtName).First());

            var result = tagProcessorMock.InnerProcessor.RepeaterTag;

            var repeaterElements = 
                TraverseUtils.ElementsBetween(result.StartRepeater, result.EndRepeater)
                             .Select(result.MakeElementCallback).ToList();
            Assert.AreEqual(1, repeaterElements.Count);

            var childrenOfFirstElement = repeaterElements.First().Elements.ToList();
            Assert.AreEqual(9, childrenOfFirstElement.Count);
            Assert.AreEqual("./Subject", childrenOfFirstElement[3].Expression);
            Assert.AreEqual(true, childrenOfFirstElement[5].IsIndex);
            Assert.AreEqual("./ExpireDate", childrenOfFirstElement[7].Expression);
            Assert.AreEqual("//test/certificates", result.Source);
        }

        [TestMethod]
        public void TestParseNested()
        {
            var parser = new RepeaterParser();
            var rootProcessor = new RootProcessor();
            parser.Parse(rootProcessor, this.nestedDocumentRoot.Descendants(WordMl.SdtName).First());

            var repeaterProcessor = rootProcessor.Processors.First();
            var result = ((RepeaterProcessor)repeaterProcessor).RepeaterTag;

            var repeaterElements = 
                TraverseUtils.ElementsBetween(result.StartRepeater, result.EndRepeater)
                             .Select(result.MakeElementCallback).ToList();
            Assert.AreEqual(2, repeaterElements.Count);

            var textTag = repeaterElements.First();

            Assert.IsTrue(textTag.XElement.IsSdt());
            Assert.IsTrue(textTag.XElement.IsTag("Text"));

            var repeaterContent = repeaterElements[1].Elements.ToList();

            Assert.AreEqual(10, repeaterContent.Count);
            var textSdt = repeaterContent[3];
            Assert.IsNull(textSdt.Expression);
            Assert.IsTrue(textSdt.XElement.IsSdt());
            Assert.IsTrue(textSdt.XElement.IsTag("Text"));
            Assert.AreEqual("./Subject", repeaterContent[4].Expression);
            Assert.AreEqual(true, repeaterContent[6].IsIndex);
            Assert.AreEqual("./ExpireDate", repeaterContent[8].Expression);
            Assert.AreEqual("//test/certificates", result.Source);

            Assert.IsTrue(repeaterProcessor.Processors.All(p => p is TextProcessor));
            Assert.AreEqual(2, repeaterProcessor.Processors.Count);
        }
    }
}
