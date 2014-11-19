using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TsSoft.Commons.Utils;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    [TestClass]
    public class HtmlContentParserTest
    {
        [TestMethod]
        public void TestParse()
        {
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "HtmlContentParserTest.xml"))
            {
                var doc = XDocument.Load(docStream);
                var processorMock = new TagProcessorMock<HtmlContentProcessor>();
                var parser = new HtmlContentParser();
                parser.Parse(processorMock, doc.Descendants(WordMl.SdtName).First());
                var processor = processorMock.InnerProcessor;
                var tag = processor.HtmlTag;
                Assert.IsNotNull(tag);
                Assert.AreEqual("//test/htmlcontent", tag.Expression);
            }
        }
    }
}
