using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    [TestClass]
    public class GeneralParserTest
    {
        private GeneralParser generalParser;
        private XDocument doc;

        [TestInitialize]
        public void Init()
        {
            this.generalParser = new GeneralParser();
        }

        [TestMethod]
        public void TestParseRepeaterNested()
        {  
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "text_in_repeater_document.xml");
            this.doc = XDocument.Load(docStream);

            var expectedRootProcessor = new RootProcessor();
            var expectedRepeaterProcessor = new RepeaterProcessor();
            expectedRepeaterProcessor.AddProcessor(new TextProcessor());
            var expectedRepeaterProcessors = new List<ITagProcessor>(expectedRepeaterProcessor.Processors);
            expectedRootProcessor.AddProcessor(expectedRepeaterProcessor);   
            var expectedRootProcessors = new List<ITagProcessor>(expectedRootProcessor.Processors);

            var actualRoot = new RootProcessor();
            this.generalParser.Parse(actualRoot, this.doc.Root);
            
            var actualRootProcessors = new List<ITagProcessor>(actualRoot.Processors);
            
            Assert.AreEqual(actualRootProcessors.Count, expectedRootProcessors.Count);
            Assert.AreEqual(expectedRootProcessors[0].GetType(), actualRootProcessors[0].GetType());
            Assert.AreEqual(expectedRootProcessors[0].Processors.Count, actualRootProcessors[0].Processors.Count);
            var actualRepeaterProcessors = new List<ITagProcessor>(actualRootProcessors[0].Processors);
            Assert.AreEqual(expectedRepeaterProcessors[0].GetType(), actualRepeaterProcessors[0].GetType());
            Assert.AreEqual(expectedRepeaterProcessors.Count, actualRepeaterProcessors.Count);
        }
    }
}
