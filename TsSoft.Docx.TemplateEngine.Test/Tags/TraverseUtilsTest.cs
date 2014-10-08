using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    using System.Linq;
    using System.Xml.Linq;

    using TsSoft.Commons.Utils;
    using TsSoft.Docx.TemplateEngine.Tags;

    [TestClass]
    public class TraverseUtilsTest
    {
        [TestMethod]
        public void TestNextElementsInOtherParagraph()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "ComplexIf.xml");
            var doc = XDocument.Load(docStream);
            var documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

            var startIf =
                documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If")); 
            var endIf =
                documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            var endIfCollection = TraverseUtils.NextTagElements(startIf, "EndIf").ToList();

            Assert.AreEqual(1, endIfCollection.Count());
            Assert.AreEqual(endIf, endIfCollection[0]);
        }
        
        [TestMethod]
        public void TestNextElementsFromParagraphToBody()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "ComplexIf.xml");
            var doc = XDocument.Load(docStream);
            var documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

            var startIf =
                documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If"));  
            var endIf =
                documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            endIf.Remove();
            documentRoot.Add(endIf);

            var endIfCollection = TraverseUtils.NextTagElements(startIf, "EndIf").ToList();

            Assert.AreEqual(1, endIfCollection.Count());
            Assert.AreEqual(endIf, endIfCollection[0]);
        } 
        
        [TestMethod]
        public void TestNextElementsFromBodyToParagraph()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "ComplexIf.xml");
            var doc = XDocument.Load(docStream);
            var documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

            var startIf =
                documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If"));  
            var endIf =
                documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            startIf.Remove();
            documentRoot.AddFirst(startIf);

            var endIfCollection = TraverseUtils.NextTagElements(startIf, "EndIf").ToList();

            Assert.AreEqual(1, endIfCollection.Count());
            Assert.AreEqual(endIf, endIfCollection[0]);
        }

        [TestMethod]
        public void TestAnySdtInOtherParagraph()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "ComplexIf.xml");
            var doc = XDocument.Load(docStream);
            var documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

            var startIf =
                documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If"));
            var endIf =
                documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            var endIfCollection = TraverseUtils.NextTagElements(startIf).ToList();

            Assert.AreEqual(1, endIfCollection.Count());
            Assert.AreEqual(endIf, endIfCollection[0]);
        }

        [TestMethod]
        public void TestAnySdtFromParagraphToBody()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "ComplexIf.xml");
            var doc = XDocument.Load(docStream);
            var documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

            var startIf =
                documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If"));
            var endIf =
                documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            endIf.Remove();
            documentRoot.Add(endIf);

            var endIfCollection = TraverseUtils.NextTagElements(startIf).ToList();

            Assert.AreEqual(1, endIfCollection.Count());
            Assert.AreEqual(endIf, endIfCollection[0]);
        }

        [TestMethod]
        public void TestAnySdtFromBodyToParagraph()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "ComplexIf.xml");
            var doc = XDocument.Load(docStream);
            var documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

            var startIf =
                documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If"));
            var endIf =
                documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            startIf.Remove();
            documentRoot.AddFirst(startIf);

            var endIfCollection = TraverseUtils.NextTagElements(startIf).ToList();

            Assert.AreEqual(1, endIfCollection.Count());
            Assert.AreEqual(endIf, endIfCollection[0]);
        }
    }
}
