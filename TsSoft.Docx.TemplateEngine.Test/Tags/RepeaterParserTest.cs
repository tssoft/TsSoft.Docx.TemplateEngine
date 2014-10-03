using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Parsers;
using TsSoft.Docx.TemplateEngine.Tags;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    [TestClass]
    public class RepeaterParserTest
    {

        private XElement documentRoot;

        [TestInitialize]
        public void Initialize()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "Repeater_Ok.xml");
            var doc = XDocument.Load(docStream);
            documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");
         
        }

        [TestMethod]
        public void TestRepeaterNotClosed()
        {
            const string tagName = "EndRepeater";
            var endRepeater = TraverseUtils.TagElement(documentRoot, tagName);
            endRepeater.Remove();
            var parser = new RepeaterParser();
            try
            {
                parser.Do(documentRoot.Descendants(WordMl.SdtName).First());
                Assert.Fail("An exception shoud've been thrown");
            }
            catch (Exception e)
            {
                Assert.AreEqual(String.Format(MessageStrings.TagNotFoundOrEmpty, tagName), e.Message);
            }
        } 
        
        [TestMethod]
        public void TestContentNotClosed()
        {
            const string tagName = "EndContent";
            var endContent = TraverseUtils.TagElement(documentRoot, tagName);
            endContent.Remove();
            var parser = new RepeaterParser();
            try
            {
                parser.Do(documentRoot.Descendants(WordMl.SdtName).First());
                Assert.Fail("An exception shoud've been thrown");
            }
            catch (Exception e)
            {
                Assert.AreEqual(String.Format(MessageStrings.TagNotFoundOrEmpty, tagName), e.Message);
            }
        }  
        
        [TestMethod]
        public void TestNoItems()
        {
            string tagName = "Items";
            var items = TraverseUtils.TagElement(documentRoot, tagName);
            items.Remove();
            var parser = new RepeaterParser();
            try
            {
                parser.Do(documentRoot.Descendants(WordMl.SdtName).First());
                Assert.Fail("An exception shoud've been thrown");
            }
            catch (Exception e)
            {
                Assert.AreEqual(String.Format(MessageStrings.TagNotFoundOrEmpty, tagName), e.Message);
            }
        }

        [TestMethod]
        public void TestOkay()
        {
            var parser = new RepeaterParser();
            var result = parser.Do(documentRoot.Descendants(WordMl.SdtName).First());

            var repeaterElements = result.Content.ToList();
            Assert.AreEqual(1, repeaterElements.Count);
            
            var childrenOfFirstElement = repeaterElements.First().Elements.ToList();
            Assert.AreEqual(9, childrenOfFirstElement.Count);
            Assert.AreEqual("./Subject", childrenOfFirstElement[3].Expression);
            Assert.AreEqual(true, childrenOfFirstElement[5].IsIndex);
            Assert.AreEqual("./ExpireDate", childrenOfFirstElement[7].Expression);
            Assert.AreEqual("//test/certificates", result.Source);
        }


    }
}
