﻿using Microsoft.VisualStudio.TestTools.UnitTesting;

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
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "IfParserTest.xml");
            var doc = XDocument.Load(docStream);
            var documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

            var parser = new IfParser();
            var startElement = TraverseUtils.TagElement(documentRoot, "If");
            var endElement = TraverseUtils.TagElement(documentRoot, "EndIf");
            var parentProcessor = new RootProcessor();
            parser.Parse(parentProcessor, startElement);

            var ifProcessor = (IfProcessor)parentProcessor.Processors.First();
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
            Assert.IsTrue(paragraphChildren[5].IsTag("Text"));
            Assert.AreEqual(startElement, ifTag.StartIf);
            Assert.AreEqual(endElement, ifTag.EndIf);

            var textProcessor = (TextProcessor)ifProcessor.Processors.First();
            Assert.IsNotNull(textProcessor);
            Assert.IsNotNull(textProcessor.TextTag);
            Assert.AreEqual("//test/text", textProcessor.TextTag.Expression);
        }

        [TestMethod]
        public void TestParseComplex()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "ComplexIf.xml");
            var doc = XDocument.Load(docStream);
            var documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

            var parser = new IfParser();
            var startElement = TraverseUtils.TagElement(documentRoot, "If");
            var endElement = TraverseUtils.TagElement(documentRoot, "EndIf");
            var parentProcessor = new RootProcessor();
            parser.Parse(parentProcessor, startElement);

            var ifProcessor = (IfProcessor)parentProcessor.Processors.First();
            var ifTag = ifProcessor.Tag;
            const string IfCondition = "//test/condition";

            Assert.IsNotNull(ifProcessor);
            Assert.AreEqual(IfCondition, ifTag.Conidition);
           
            Assert.AreEqual(startElement, ifTag.StartIf);
            Assert.AreEqual(endElement, ifTag.EndIf);

            var content = ifTag.IfContent.ToList();

            Assert.AreEqual(18, content.Count);
            Assert.IsTrue(content.Take(7).All(e => e.Name.Equals(WordMl.TextRunName)));
            Assert.IsTrue(content.Skip(7).Take(1).All(e => e.Name.Equals(WordMl.ParagraphName)));
            Assert.IsTrue(content.Skip(8).Take(10).All(e => e.Name.Equals(WordMl.TextRunName)));
        }
    }
}
