using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    [TestClass]
    public class GeneralParserTest
    {

        private GeneralParser generalParser;       

        [TestInitialize]
        public void Init()
        {
            generalParser = new GeneralParser();
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "GeneralParserTest.xml");
            var doc = XDocument.Load(docStream);
        }
        [TestMethod]
        public void TestParse()
        {
            var tagProcessorMock = new Mock<ITagProcessor>();
            //generalParser.Parse(tagProcessorMock.Object,);
            
        }

    }
}
