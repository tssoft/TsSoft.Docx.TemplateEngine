using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Test.Tags.Processors
{
    [TestClass]
    public class HtmlContentProcessorTest
    {
        [TestMethod]
        public void TestProcess()
        {
            const string HtmlEscapedString = @"
                        &lt;html&gt;
                        &lt;head/&gt;
                        &lt;body&gt;
                        &lt;p&gt;Html test&lt;/p&gt;
                        &lt;/body&gt;
                        &lt;/html&gt;
                        ";
            const string path = "//test/htmlcontent";
            var data = new XElement("test", new XElement("htmlcontent", HtmlEscapedString));
            var htmlContentTag = new XElement(
                WordMl.SdtName,
                new XElement(
                    WordMl.SdtPrName, new XElement(
                                          WordMl.TagName,
                                          new XAttribute(WordMl.ValAttributeName, "htmlcontent"))),
                new XElement(
                    WordMl.SdtContentName, path));
            var document = new XElement("body", htmlContentTag);
            var processor = new HtmlContentProcessor()
                {
                    DataReader = new DataReader(data),
                    HtmlTag = new HtmlContentTag()
                        {
                            Expression = path,
                            TagNode = htmlContentTag
                        }
                };
            processor.Process();
            Assert.AreEqual(HtmlContentProcessor.ProcessedHtmlContentTagName,
                            htmlContentTag.Element(WordMl.SdtPrName)
                                          .Element(WordMl.TagName)
                                          .Attribute(WordMl.ValAttributeName)
                                          .Value);
            Assert.AreNotEqual(path, htmlContentTag.Element(WordMl.SdtContentName).Value);
        }
    }
}
