using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    using System.Linq;
    using System.Xml.Linq;

    using TsSoft.Commons.Utils;
    using TsSoft.Docx.TemplateEngine.Tags;
    using TsSoft.Docx.TemplateEngine.Tags.Processors;

    [TestClass]
    public class IfParserTest
    {
        [TestMethod]
        public void TestParse()
        {
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "IfParserTest.xml"))
            {
                var doc = XDocument.Load(docStream);
                var documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

                var parser = new IfParser();
                var startElement = TraverseUtils.TagElement(documentRoot, "If");
                var endElement = TraverseUtils.TagElement(documentRoot, "EndIf");
                var parentProcessor = new RootProcessor();
                parser.Parse(parentProcessor, startElement);

                var ifProcessor = (IfProcessor) parentProcessor.Processors.First();
                var ifTag = ifProcessor.Tag;
                const string IfCondition = "//test/condition";

                Assert.IsNotNull(ifProcessor);
                Assert.AreEqual(IfCondition, ifTag.Conidition);
                var content = ifTag.IfContent.ToList();
                Assert.AreEqual(1, content.Count);
                Assert.AreEqual(WordMl.ParagraphName, content[0].Name);
                var paragraphChildren = content[0].Elements().ToList();
                Assert.AreEqual(6, paragraphChildren.Count());
                Assert.AreEqual(WordMl.ParagraphPropertiesName, paragraphChildren[0].Name);
                Assert.AreEqual(WordMl.ProofingErrorAnchorName, paragraphChildren[1].Name);
                Assert.AreEqual(WordMl.TextRunName, paragraphChildren[2].Name);
                Assert.AreEqual("Hello", paragraphChildren[2].Value);
                Assert.AreEqual(WordMl.ProofingErrorAnchorName, paragraphChildren[3].Name);
                Assert.AreEqual(WordMl.TextRunName, paragraphChildren[4].Name);
                Assert.AreEqual(", World!", paragraphChildren[4].Value);
                Assert.AreEqual(WordMl.SdtName, paragraphChildren[5].Name);
                Assert.IsTrue(paragraphChildren[5].IsTag("text"));
                Assert.AreEqual(startElement, ifTag.StartIf);
                Assert.AreEqual(endElement, ifTag.EndIf);

                var textProcessor = (TextProcessor) ifProcessor.Processors.FirstOrDefault();
                Assert.IsNotNull(textProcessor);
                Assert.IsNotNull(textProcessor.TextTag);
                Assert.AreEqual("//test/text", textProcessor.TextTag.Expression);
            }
        }

        [TestMethod]
        public void TestParseComplex()
        {
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "ComplexIf.xml"))
            {
                var doc = XDocument.Load(docStream);
                var documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

                var parser = new IfParser();
                var startElement = TraverseUtils.TagElement(documentRoot, "If");
                var endElement = TraverseUtils.TagElement(documentRoot, "EndIf");
                var parentProcessor = new RootProcessor();
                parser.Parse(parentProcessor, startElement);

                var ifProcessor = (IfProcessor) parentProcessor.Processors.First();
                var ifTag = ifProcessor.Tag;
                const string IfCondition = "//test/condition";

                Assert.IsNotNull(ifProcessor);
                Assert.AreEqual(IfCondition, ifTag.Conidition);

                Assert.AreEqual(startElement, ifTag.StartIf);
                Assert.AreEqual(endElement, ifTag.EndIf);

                var elements = ifTag.IfContent.ToList();

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

                Assert.AreEqual(startElement, ifTag.StartIf);
                Assert.AreEqual(endElement, ifTag.EndIf);
            }
        }
        [TestMethod]
        public void TestParseCplx()
        {
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "document_dontworking_if.xml"))
            {
                var doc = XDocument.Load(docStream);
                var documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

                var parser = new IfParser();
                var startElement = TraverseUtils.TagElement(documentRoot, "If");
                var endElement = TraverseUtils.TagElement(documentRoot, "EndIf");
                var parentProcessor = new RootProcessor();
                parser.Parse(parentProcessor, startElement);

                var ifProcessor = (IfProcessor) parentProcessor.Processors.First();
                var ifTag = ifProcessor.Tag;
                const string IfCondition = "//test/condition";

                Assert.IsNotNull(ifProcessor);

                //Assert.IsTrue(IfCondition.Equals(ifTag.Conidition));
                CollectionAssert.AreEqual(Encoding.UTF8.GetBytes(IfCondition), Encoding.UTF8.GetBytes(ifTag.Conidition));

                Assert.AreEqual(startElement, ifTag.StartIf);
                Assert.AreEqual(endElement, ifTag.EndIf);

                var elements = ifTag.IfContent.ToList();

                Assert.AreEqual(1, elements.Count());
                /*
                Assert.IsTrue(elements.Take(6).All(e => e.Name.Equals(WordMl.TextRunName)));
                Assert.IsTrue(elements.Skip(6).Take(1).All(e => e.Name.Equals(WordMl.ParagraphName)));
                Assert.IsTrue(elements.Skip(7).Take(1).All(e => e.Name.Equals(WordMl.ParagraphPropertiesName)));
                Assert.IsTrue(elements.Skip(8).Take(1).All(e => e.Name.Equals(WordMl.BookmarkStartName)));
                Assert.IsTrue(elements.Skip(9).Take(4).All(e => e.Name.Equals(WordMl.TextRunName)));
                Assert.IsTrue(elements.Skip(13).Take(1).All(e => e.Name.Equals(WordMl.ProofingErrorAnchorName)));
                Assert.IsTrue(elements.Skip(14).Take(1).All(e => e.Name.Equals(WordMl.TextRunName)));
                Assert.IsTrue(elements.Skip(15).Take(1).All(e => e.Name.Equals(WordMl.ProofingErrorAnchorName)));
                Assert.IsTrue(elements.Skip(16).Take(1).All(e => e.Name.Equals(WordMl.TextRunName)));
                */
                Assert.AreEqual(startElement, ifTag.StartIf);
                Assert.AreEqual(endElement, ifTag.EndIf);
            }
        }

    }
}
