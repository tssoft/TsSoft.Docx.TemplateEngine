using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Tags;
namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    [TestClass]
    public class TraverseUtilsTest
    {
        private XElement documentRoot;

        [TestInitialize]
        public void TestInit()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "ComplexIf.xml");
            var doc = XDocument.Load(docStream);
            this.documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");
        }

        [TestMethod]
        public void TestNextElementsInOtherParagraph()
        {
            var startIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If"));
            var endIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            var endIfCollection = TraverseUtils.NextTagElements(startIf, "EndIf").ToList();

            Assert.AreEqual(1, endIfCollection.Count());
            Assert.AreEqual(endIf, endIfCollection[0]);
        }

        [TestMethod]
        public void TestNextElementsFromParagraphToBody()
        {
            var startIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If"));  
            var endIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            endIf.Remove();
            this.documentRoot.Add(endIf);

            var endIfCollection = TraverseUtils.NextTagElements(startIf, "EndIf").ToList();

            Assert.AreEqual(1, endIfCollection.Count());
            Assert.AreEqual(endIf, endIfCollection[0]);
        }

        [TestMethod]
        public void TestNextElementsFromBodyToParagraph()
        {
            var startIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If"));  
            var endIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            startIf.Remove();
            this.documentRoot.AddFirst(startIf);

            var endIfCollection = TraverseUtils.NextTagElements(startIf, "EndIf").ToList();

            Assert.AreEqual(1, endIfCollection.Count());
            Assert.AreEqual(endIf, endIfCollection[0]);
        }

        [TestMethod]
        public void TestAnySdtInOtherParagraph()
        {
            var startIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If"));
            var endIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            var endIfCollection = TraverseUtils.NextTagElements(startIf).ToList();

            Assert.AreEqual(1, endIfCollection.Count());
            Assert.AreEqual(endIf, endIfCollection[0]);
        }

        [TestMethod]
        public void TestAnySdtFromParagraphToBody()
        {
            var startIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If"));
            var endIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            endIf.Remove();
            this.documentRoot.Add(endIf);

            var endIfCollection = TraverseUtils.NextTagElements(startIf).ToList();

            Assert.AreEqual(1, endIfCollection.Count());
            Assert.AreEqual(endIf, endIfCollection[0]);
        }

        [TestMethod]
        public void TestAnySdtFromBodyToParagraph()
        {
            var startIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If"));
            var endIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            startIf.Remove();
            this.documentRoot.AddFirst(startIf);

            var endIfCollection = TraverseUtils.NextTagElements(startIf).ToList();

            Assert.AreEqual(1, endIfCollection.Count());
            Assert.AreEqual(endIf, endIfCollection[0]);
        }

        [TestMethod]
        public void TestElementBetweenEachInDifferentParagraphs()
        {
            var startIf =
              this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "If"));
            var endIf =
                this.documentRoot.Descendants(WordMl.SdtName).Single(d => d.Descendants(WordMl.TagName).Any(t => t.Attribute(WordMl.ValAttributeName).Value == "EndIf"));

            var elements = TraverseUtils.ElementsBetween(startIf, endIf).ToList();

            Assert.AreEqual(17, elements.Count());
            Assert.IsTrue(elements.Take(6).All(e => e.Name.Equals(WordMl.TextRunName)));
            Assert.IsTrue(elements.Skip(6).Take(1).All(e => e.Name.Equals(WordMl.ParagraphName)));
            Assert.IsTrue(elements.Skip(7).Take(1).All(e => e.Name.Equals(WordMl.ParagraphPropertiesName)));
            Assert.IsTrue(elements.Skip(8).Take(1).All(e => e.Name.Equals(WordMl.BookmarkStartName)));
            Assert.IsTrue(elements.Skip(9).Take(4).All(e => e.Name.Equals(WordMl.TextRunName)));
            Assert.IsTrue(elements.Skip(13).Take(1).All(e => e.Name.Equals(WordMl.ProofingErrorAnchorName)));
            Assert.IsTrue(elements.Skip(14).Take(1).All(e => e.Name.Equals(WordMl.TextRunName)));
            Assert.IsTrue(elements.Skip(15).Take(1).All(e => e.Name.Equals(WordMl.ProofingErrorAnchorName)));
            Assert.IsTrue(elements.Skip(16).Take(1).All(e => e.Name.Equals(WordMl.TextRunName)));
        }
    }
}
