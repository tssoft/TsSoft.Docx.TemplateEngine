using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Parsers;
using TsSoft.Docx.TemplateEngine.Tags;

namespace TsSoft.Docx.TemplateEngine.Test.Parsers
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
            var endRepeater = TraverseUtils.TagElement(documentRoot, "EndRepeater");
            endRepeater.Remove();
            var parser = new RepeaterParser();
            try
            {
                parser.Do(documentRoot.Descendants(WordMl.SdtName).First());
                Assert.Fail("An exception shoud've been thrown");
            }
            catch (Exception e)
            {
                Assert.AreEqual("Required tag EndRepeater not found", e.Message);
            }
        } 
        
        [TestMethod]
        public void TestContentNotClosed()
        {
            var endContent = TraverseUtils.TagElement(documentRoot, "EndContent");
            endContent.Remove();
            var parser = new RepeaterParser();
            try
            {
                parser.Do(documentRoot.Descendants(WordMl.SdtName).First());
                Assert.Fail("An exception shoud've been thrown");
            }
            catch (Exception e)
            {
                Assert.AreEqual("Required tag EndContent not found", e.Message);
            }
        }  
        
        [TestMethod]
        public void TestNoItems()
        {
            var items = TraverseUtils.TagElement(documentRoot, "Items");
            items.Remove();
            var parser = new RepeaterParser();
            try
            {
                parser.Do(documentRoot.Descendants(WordMl.SdtName).First());
                Assert.Fail("An exception shoud've been thrown");
            }
            catch (Exception e)
            {
                Assert.AreEqual("Required tag Items not found", e.Message);
            }
        }

        [TestMethod]
        public void TestOkay()
        {
            var parser = new RepeaterParser();
            RepeaterTag result = parser.Do(documentRoot.Descendants(WordMl.SdtName).First());

            List<RepeaterElement> repeaterElements = result.Content.ToList();
            Assert.AreEqual(1, repeaterElements.Count);
            
            List<RepeaterElement> childrenOfFirstElement = repeaterElements.First().Elements.ToList();
            Assert.AreEqual(9, childrenOfFirstElement.Count);
            Assert.AreEqual("./Subject", childrenOfFirstElement[3].Expression);
            Assert.AreEqual(true, childrenOfFirstElement[5].IsIndex);
            Assert.AreEqual("./ExpireDate", childrenOfFirstElement[7].Expression);
            Assert.AreEqual("//test/certificates", result.Source);
        }


    }
}
