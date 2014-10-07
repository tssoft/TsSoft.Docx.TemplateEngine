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
        /*
        private Mock<ITagProcessor> tagProcessorMock;
        private Mock<TextParser> textParserMock;
        private Mock<TableParser> tableParserMock;
        private Mock<RepeaterParser> repeaterParserMock;

        private Mock<RootProcessor> rootProcessorMock;
       */
        /*
        private void AssertingTree(IList<ITagProcessor> expected, IList<ITagProcessor> actual)
        {
            Assert.AreEqual(expected.Count, actual.Count);
            for (int i = 0; i < expected.Count; i++)
            {                
                Assert.AreEqual(expected[i].GetType(), actual[i].GetType());
               // Assert.AreEqual(expected[i].);
                AssertingTree(new List<ITagProcessor>(actual[i].Processors),new List<ITagProcessor>(expected[i].Processors));
            }
        }
         * */

       

        [TestInitialize]
        public void Init()
        {
            generalParser = new GeneralParser();            
        }

        [TestMethod]
        public void TestParseTableNested()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "if_in_table_document.xml");
            doc = XDocument.Load(docStream);

            var expectedRootProcessor = new RootProcessor();
            ITagProcessor expectedTableProcessor = new TableProcessor();
            expectedTableProcessor.AddProcessor(new IfProcessor());
            IList<ITagProcessor> expectedTableProcessors = new List<ITagProcessor>(expectedTableProcessor.Processors);
            expectedRootProcessor.AddProcessor(expectedRootProcessor);
            
            IList<ITagProcessor> expectedRootProcessors = new List<ITagProcessor>(expectedRootProcessor.Processors);
            
            ITagProcessor actualRootProcessor = new RootProcessor();
            generalParser.Parse(actualRootProcessor, doc.Root);
            IList<ITagProcessor> actualRootProcessors = new List<ITagProcessor>();

            Assert.AreEqual(expectedRootProcessors.Count, actualRootProcessors.Count);
            Assert.AreEqual(expectedRootProcessors[0].GetType(), actualRootProcessors[0].GetType());
            

        }


        [TestMethod]
        public void TestParseIfNested()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "text_in_if_document.xml");
            doc = XDocument.Load(docStream);
            
            var expectedRootProcessor = new RootProcessor();
            ITagProcessor expectedIfProcessor = new IfProcessor();
            
            expectedIfProcessor.AddProcessor(new TextProcessor());
            IList<ITagProcessor> expectedIfProcessors = new List<ITagProcessor>(expectedIfProcessor.Processors);
            expectedRootProcessor.AddProcessor(expectedIfProcessor);            
            IList<ITagProcessor> expectedRootProcessors = new List<ITagProcessor>(expectedRootProcessor.Processors);

            ITagProcessor actualRootProcessor = new RootProcessor();
            generalParser.Parse(actualRootProcessor, doc.Root);

            IList<ITagProcessor> actualRootProcessors = new List<ITagProcessor>(actualRootProcessor.Processors);
            Assert.AreEqual(expectedRootProcessors.Count, actualRootProcessors.Count);
            Assert.AreEqual(expectedRootProcessors[0].GetType(),actualRootProcessors[0].GetType());
            IList<ITagProcessor> actualIfProcessors = new List<ITagProcessor>(actualRootProcessors[0].Processors);
            Assert.AreEqual(expectedIfProcessors.Count, actualIfProcessors.Count);
            Assert.AreEqual(expectedIfProcessors[0].GetType(), actualIfProcessors[0].GetType());
            Assert.AreEqual(expectedIfProcessors[0].Processors.Count, expectedIfProcessors[0].Processors.Count); 
             
        }

        [TestMethod]
        public void TestParseRepeaterNested()
        {  

            var docStream = AssemblyResourceHelper.GetResourceStream(this, "text_in_repeater_document.xml");
            doc = XDocument.Load(docStream);

            var expectedRootProcessor = new RootProcessor();

            ITagProcessor expectedRepeaterProcessor = new RepeaterProcessor();
            expectedRepeaterProcessor.AddProcessor(new TextProcessor());
            IList<ITagProcessor> expectedRepeaterProcessors = new List<ITagProcessor>(expectedRepeaterProcessor.Processors);
            expectedRootProcessor.AddProcessor(expectedRepeaterProcessor);               
            IList<ITagProcessor> expectedRootProcessors = new List<ITagProcessor>(expectedRootProcessor.Processors);

            ITagProcessor actualRoot = new RootProcessor();
            generalParser.Parse(actualRoot, doc.Root);            
            IList<ITagProcessor> actualRootProcessors = new List<ITagProcessor>(actualRoot.Processors);            
            Assert.AreEqual(actualRootProcessors.Count, expectedRootProcessors.Count);
            Assert.AreEqual(expectedRootProcessors[0].GetType(), actualRootProcessors[0].GetType()); 
            Assert.AreEqual(expectedRootProcessors[0].Processors.Count, actualRootProcessors[0].Processors.Count);
            IList<ITagProcessor> actualRepeaterProcessors = new List<ITagProcessor>(actualRootProcessors[0].Processors);         
            Assert.AreEqual(expectedRepeaterProcessors[0].GetType(), actualRepeaterProcessors[0].GetType());
            Assert.AreEqual(expectedRepeaterProcessors.Count, actualRepeaterProcessors.Count);

           
        }

    }
}
