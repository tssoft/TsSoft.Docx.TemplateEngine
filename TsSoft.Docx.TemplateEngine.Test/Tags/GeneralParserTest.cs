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
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "GeneralParserTest.xml"))
            {
                this.doc = XDocument.Load(docStream);
            }
        }

        [TestMethod]
        public void TestParse()
        {
            var expectedTableProcessor = new TableProcessor();
            var expectedIfProcessor = new IfProcessor();
            expectedIfProcessor.AddProcessor(new TextProcessor());
            expectedTableProcessor.AddProcessor(new TextProcessor());
            expectedTableProcessor.AddProcessor(expectedIfProcessor);
            var expectedRepeaterProcessor = new RepeaterProcessor();
            expectedRepeaterProcessor.AddProcessor(new TextProcessor());
            var expectedRootProcessor = new RootProcessor();
            expectedRootProcessor.AddProcessor(expectedTableProcessor);
            expectedRootProcessor.AddProcessor(expectedRepeaterProcessor);

            var actualRootProcessor = new RootProcessor();
            generalParser.Parse(actualRootProcessor, this.doc.Root);

            Assert.AreEqual(2, actualRootProcessor.Processors.Count);
            var actualInnerRootProcessors = new List<ITagProcessor>(actualRootProcessor.Processors);
            Assert.AreEqual(typeof(TableProcessor), actualInnerRootProcessors[0].GetType());
            Assert.AreEqual(typeof(RepeaterProcessor), actualInnerRootProcessors[1].GetType());

            var actualTableProcessor = actualInnerRootProcessors[0];
            Assert.AreEqual(2, actualTableProcessor.Processors.Count);
            var actualInnerTableProcessors = new List<ITagProcessor>(actualTableProcessor.Processors);
            Assert.AreEqual(typeof(TextProcessor), actualInnerTableProcessors[0].GetType());
            Assert.AreEqual(typeof(IfProcessor), actualInnerTableProcessors[1].GetType());

            Assert.AreEqual(0, actualInnerTableProcessors[0].Processors.Count);
            var actualIfProcessor = actualInnerTableProcessors[1];

            Assert.AreEqual(1, actualIfProcessor.Processors.Count);
            var actualInnerIfProcessors = new List<ITagProcessor>(actualIfProcessor.Processors);
            Assert.AreEqual(typeof(TextProcessor), actualInnerIfProcessors[0].GetType());

            var actualRepeaterProcessor = actualInnerRootProcessors[1];
            Assert.AreEqual(1,actualRepeaterProcessor.Processors.Count);
            var actualInnerRepeaterProcessors = new List<ITagProcessor>(actualRepeaterProcessor.Processors);
            Assert.AreEqual(typeof(TextProcessor), actualInnerRepeaterProcessors[0].GetType());
            Assert.AreEqual(0, actualInnerRepeaterProcessors[0].Processors.Count);
        }


        [TestMethod]
        public void TestParseRepeaterInIf()
        {
            //var docStream = AssemblyResourceHelper.GetResourceStream(this, "")
            
        }

        [TestMethod]
        public void TestParseIfNested()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "text_in_if_document.xml");
            doc = XDocument.Load(docStream);
            
            var expectedRootProcessor = new RootProcessor();
            ITagProcessor expectedIfProcessor = new IfProcessor();
            
            expectedIfProcessor.AddProcessor(new TextProcessor());
            var expectedIfProcessors = new List<ITagProcessor>(expectedIfProcessor.Processors);
            expectedRootProcessor.AddProcessor(expectedIfProcessor);            
            var expectedRootProcessors = new List<ITagProcessor>(expectedRootProcessor.Processors);

            var actualRootProcessor = new RootProcessor();
            generalParser.Parse(actualRootProcessor, doc.Root);

            var actualRootProcessors = new List<ITagProcessor>(actualRootProcessor.Processors);
            Assert.AreEqual(expectedRootProcessors.Count, actualRootProcessors.Count);
            Assert.AreEqual(expectedRootProcessors[0].GetType(),actualRootProcessors[0].GetType());
            var actualIfProcessors = new List<ITagProcessor>(actualRootProcessors[0].Processors);
            Assert.AreEqual(expectedIfProcessors.Count, actualIfProcessors.Count);
            Assert.AreEqual(expectedIfProcessors[0].GetType(), actualIfProcessors[0].GetType());
            Assert.AreEqual(expectedIfProcessors[0].Processors.Count, expectedIfProcessors[0].Processors.Count); 
             
        }

        [TestMethod]
        public void TestParseTableNested()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "if_in_table_document.xml");
            doc = XDocument.Load(docStream);

            var expectedRootProcessor = new RootProcessor();
            var expectedTableProcessor = new TableProcessor();
            expectedTableProcessor.AddProcessor(new IfProcessor());
            expectedRootProcessor.AddProcessor(expectedTableProcessor);
            var expectedTableProcessors = new List<ITagProcessor>(expectedTableProcessor.Processors);            
            var expectedRootProcessors = new List<ITagProcessor>(expectedRootProcessor.Processors);

            var actualRootProcessor = new RootProcessor();
            generalParser.Parse(actualRootProcessor, doc.Root);
            var actualRootProcessors = new List<ITagProcessor>(actualRootProcessor.Processors);

            Assert.AreEqual(expectedRootProcessors.Count, actualRootProcessors.Count);
            Assert.AreEqual(expectedRootProcessors[0].GetType(), actualRootProcessors[0].GetType());
            var actualTableProcessors = new List<ITagProcessor>(expectedRootProcessors[0].Processors);
            Assert.AreEqual(expectedTableProcessors.Count, actualTableProcessors.Count);
            Assert.AreEqual(expectedTableProcessors[0].GetType(), actualTableProcessors[0].GetType());
            Assert.AreEqual(expectedTableProcessors[0].Processors.Count, actualTableProcessors[0].Processors.Count);

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
