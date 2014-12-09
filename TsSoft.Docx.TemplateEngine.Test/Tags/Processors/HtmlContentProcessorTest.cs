using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
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

        [TestMethod]
        public void TestMakeHtmlContentProcessed()
        {
            const string HtmlContent = @"&lt;p&gt;Html test&lt;/p&gt;";
            const string HtmlEscapedString = @"
                        &lt;html&gt;
                        &lt;head/&gt;
                        &lt;body&gt;
                        &lt;p&gt;Html test&lt;/p&gt;
                        &lt;/body&gt;
                        &lt;/html&gt;
                        ";
            var htmlContentTag = new XElement(
                WordMl.SdtName,
                new XElement(
                    WordMl.SdtPrName, new XElement(
                                          WordMl.TagName,
                                          new XAttribute(WordMl.ValAttributeName, "htmlcontent"))),
                new XElement(
                    WordMl.SdtContentName, HtmlEscapedString));
            var body = new XElement(WordMl.BodyName, htmlContentTag);
            var actualProcessedElement = HtmlContentProcessor.MakeHtmlContentProcessed(htmlContentTag, htmlContentTag.GetExpression());
            Assert.IsNotNull(actualProcessedElement);
            Assert.AreEqual(HtmlContentProcessor.ProcessedHtmlContentTagName,
                            actualProcessedElement.Element(WordMl.SdtPrName)
                                                  .Element(WordMl.TagName)
                                                  .Attribute(WordMl.ValAttributeName)
                .Value);
            Assert.AreEqual(HtmlEscapedString, HtmlEscapedString);


        }

        [TestMethod]
        public void TestMakeHtmlContentProcessedInnerTableCell()
        {
            const string HtmlContent = @"&lt;p&gt;Html test&lt;/p&gt;";
            const string HtmlEscapedString = @"
                        &lt;html&gt;
                        &lt;head/&gt;
                        &lt;body&gt;
                        &lt;p&gt;Html test&lt;/p&gt;
                        &lt;/body&gt;
                        &lt;/html&gt;
                        ";            
            var tablecell = new XElement(WordMl.TableCellName, new XElement(WordMl.ParagraphName, "//test/path"));
            var htmlContentTag = new XElement(
                WordMl.SdtName,
                new XElement(
                    WordMl.SdtPrName, new XElement(
                                          WordMl.TagName,
                                          new XAttribute(WordMl.ValAttributeName, "htmlcontent"))),
                new XElement(
                    WordMl.SdtContentName, tablecell));
            var document = new XElement(WordMl.BodyName, htmlContentTag);
            var actualProcessedElement = HtmlContentProcessor.MakeHtmlContentProcessed(htmlContentTag, HtmlEscapedString);
            Assert.AreEqual(tablecell.Name, actualProcessedElement.Name);
            Assert.IsNotNull(actualProcessedElement.Element(WordMl.SdtName));
            Assert.AreEqual(HtmlContentProcessor.ProcessedHtmlContentTagName,
                            actualProcessedElement.Element(WordMl.SdtName)
                                                  .Element(WordMl.SdtPrName)
                                                  .Element(WordMl.TagName)
                .Attribute(WordMl.ValAttributeName)
                .Value);
            Assert.AreEqual(HttpUtility.HtmlDecode(HtmlEscapedString), actualProcessedElement.Element(WordMl.SdtName).Element(WordMl.SdtContentName).Value);
        }

        [TestMethod]
        public void TestGenerateAltChunks()
        {
            const string ExpectedHtmlString = "<html>altchunks test</html>";
            var htmlProcessedElement = new XElement(WordMl.SdtName,
                                                    new XElement(WordMl.SdtPrName,
                                                                 new XElement(WordMl.TagName,
                                                                              new XAttribute(WordMl.ValAttributeName,
                                                                                             HtmlContentProcessor
                                                                                                 .ProcessedHtmlContentTagName))),
                                                    new XElement(WordMl.SdtContentName, ExpectedHtmlString));
            var root = new XElement(WordMl.BodyName, htmlProcessedElement);
                        
            var actualGeneratedChunks = HtmlContentProcessor.GenerateAltChunks(root);
            var actualAltChunkElement = root.Elements(WordMl.AltChunkName).SingleOrDefault();
            Assert.IsNotNull(actualAltChunkElement);

            Assert.IsNotNull(actualGeneratedChunks);            
            Assert.AreEqual(ExpectedHtmlString, actualGeneratedChunks[0]);
            Assert.AreEqual(1, actualGeneratedChunks.Count);
            Assert.IsFalse(root.Descendants().Any(el => el.IsSdt()));
        }

        [TestMethod]
        public void TestGenerateAltChunksEqualHtml()
        {
            const string ExpectedHtmlString = "<html>altchunks test</html>";
            var htmlProcessedElement = new XElement(WordMl.SdtName,
                                                    new XElement(WordMl.SdtPrName,
                                                                 new XElement(WordMl.TagName,
                                                                              new XAttribute(WordMl.ValAttributeName,
                                                                                             HtmlContentProcessor
                                                                                                 .ProcessedHtmlContentTagName))),
                                                    new XElement(WordMl.SdtContentName, ExpectedHtmlString));
            var root = new XElement(WordMl.BodyName, htmlProcessedElement, htmlProcessedElement);

            var actualGeneratedChunks = HtmlContentProcessor.GenerateAltChunks(root);
            Assert.IsNotNull(actualGeneratedChunks);
            var actualAltChunkElements =
                root.Elements(WordMl.AltChunkName);
            Assert.IsNotNull(actualAltChunkElements);
            Assert.AreEqual(2, actualAltChunkElements.Count());
            Assert.AreEqual(ExpectedHtmlString, actualGeneratedChunks[0]);
            Assert.AreEqual(1, actualGeneratedChunks.Count);
            Assert.IsFalse(root.Descendants().Any(el => el.IsSdt()));
        }

        [TestMethod]
        public void TestGenerateAltChunksDeletingParagraph()
        {
            const string ExpectedHtmlString = "<html>altchunks test</html>";
            var htmlProcessedElement = new XElement(WordMl.SdtName,
                                                    new XElement(WordMl.SdtPrName,
                                                                 new XElement(WordMl.TagName,
                                                                              new XAttribute(WordMl.ValAttributeName,
                                                                                             HtmlContentProcessor
                                                                                                 .ProcessedHtmlContentTagName))),
                                                    new XElement(WordMl.SdtContentName, ExpectedHtmlString));
            var paragraph = new XElement(WordMl.ParagraphName, htmlProcessedElement);
            var root = new XElement(WordMl.BodyName, paragraph);
            var actualGeneratedChunks = HtmlContentProcessor.GenerateAltChunks(root);
            Assert.IsNotNull(actualGeneratedChunks);
            Assert.AreEqual(ExpectedHtmlString, actualGeneratedChunks[0]);
            Assert.AreEqual(1, actualGeneratedChunks.Count);
            Assert.IsFalse(root.Descendants().Any(el => el.IsSdt()));
            Assert.IsFalse(root.Descendants().Any(el => el.Name.Equals(WordMl.ParagraphName)));
        }

        [TestMethod]
        public void TestGenerateAltChunksAddingEmptyParagraphInTableCell()
        {
            const string ExpectedHtmlString = "<html>altchunks table cell test</html>";
            const string rsidR = "00FFFFFF";
            var htmlProcessedElement = new XElement(WordMl.SdtName,
                                                    new XElement(WordMl.SdtPrName,
                                                                 new XElement(WordMl.TagName,
                                                                              new XAttribute(WordMl.ValAttributeName,
                                                                                             HtmlContentProcessor
                                                                                                 .ProcessedHtmlContentTagName))),
                                                    new XElement(WordMl.SdtContentName, ExpectedHtmlString));
            var tableCell = new XElement(WordMl.TableCellName, htmlProcessedElement);
            var tableRow = new XElement(WordMl.TableRowName, new XAttribute(WordMl.RsidRPropertiesName, rsidR), tableCell);
            var root = new XElement(WordMl.BodyName, new XElement(WordMl.TableName, tableRow));
            var actualGeneratedChunks = HtmlContentProcessor.GenerateAltChunks(root);
            var actualAltChunkElement =
                root.Element(WordMl.TableName)
                    .Element(WordMl.TableRowName)
                    .Element(WordMl.TableCellName)
                    .Elements()
                    .SingleOrDefault(el => el.Name.Equals(WordMl.AltChunkName));
            Assert.IsNotNull(actualAltChunkElement);
            Assert.IsNotNull(actualGeneratedChunks);
            Assert.AreEqual(ExpectedHtmlString, actualGeneratedChunks[0]);
            Assert.AreEqual(1, actualGeneratedChunks.Count);
            Assert.IsFalse(root.Descendants().Any(el => el.IsSdt()));
            var actualLastElement =
                root.Element(WordMl.TableName)
                    .Element(WordMl.TableRowName)
                    .Element(WordMl.TableCellName)
                    .Elements()
                    .Last();
            Assert.IsNotNull(actualLastElement);
            Assert.AreEqual(WordMl.ParagraphName, actualLastElement.Name);

        }
    }
}
