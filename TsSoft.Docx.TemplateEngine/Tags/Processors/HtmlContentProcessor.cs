using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Web;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    internal class HtmlContentProcessor : AbstractProcessor
    {
        public static string ProcessedHtmlContentTagName = "processedhtmlcontent";

        public HtmlContentTag HtmlTag { get; set; }

        public static XElement MakeHtmlContentProcessed(XElement srcHtmlContentElement, string htmlContent, bool generateNewElement = false)
        {            
            var htmlContentElement = generateNewElement ? new XElement(srcHtmlContentElement) : srcHtmlContentElement;
            htmlContentElement.Element(WordMl.SdtPrName)
                .Element(WordMl.TagName)
                .Attribute(WordMl.ValAttributeName)
                .SetValue(ProcessedHtmlContentTagName);
            htmlContent = HttpUtility.HtmlDecode(htmlContent);         
            var tableCellElement = htmlContentElement.Descendants(WordMl.TableCellName).SingleOrDefault();
            htmlContentElement.Element(WordMl.SdtContentName).Value = htmlContent;            
            if (tableCellElement != null)
            {
                if (tableCellElement.Elements(WordMl.ParagraphName).Any())
                {
                    tableCellElement.Element(WordMl.ParagraphName).Remove();
                }
                tableCellElement.Add(htmlContentElement);
                htmlContentElement.AddAfterSelf(tableCellElement);
                htmlContentElement.Remove();
                return tableCellElement;
            }                       
            return htmlContentElement;
        }
        
        public static IList<string> GenerateAltChunks(XElement root)
        {
            var htmlTags = root.Descendants()
                .Where(el => el.IsTag(ProcessedHtmlContentTagName))
                .ToList();
            var htmlChunks = new List<string>();
            foreach (var htmlTag in htmlTags)
            {
                var htmlString = htmlTag.GetExpression();
                var htmlChunkIndex = htmlChunks.FindIndex(html => html.Equals(htmlString));
                if (htmlChunkIndex == -1)
                {
                    htmlChunks.Add(htmlString);
                    htmlChunkIndex = htmlChunks.Count - 1;
                }
                var parent = htmlTag.Parent;
                var altChunkElement = DocxHelper.CreateAltChunkElement(htmlChunkIndex + 1);
                if (!parent.Name.Equals(WordMl.ParagraphName))
                {
                    htmlTag.AddAfterSelf(altChunkElement);
                }
                else
                {
                    parent.AddAfterSelf(altChunkElement);
                    if (parent.Elements().Count(el => !el.Name.Equals(WordMl.ParagraphPropertiesName)) == 1)
                    {
                        parent.Remove();
                    }
                }
                if (altChunkElement.Parent.Name.Equals(WordMl.TableCellName) && (altChunkElement.NextElement() == null))
                {
                    DocxHelper.AddEmptyParagraphInTableCell(altChunkElement);
                }
                htmlTag.Remove();
            }
            return htmlChunks;
        }
        
        public override void Process()
        {            
            base.Process();
            var element = this.HtmlTag.TagNode;
            MakeHtmlContentProcessed(element, DataReader.ReadText(this.HtmlTag.Expression));
        }
    }
}
