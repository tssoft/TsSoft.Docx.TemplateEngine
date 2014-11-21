using System;
using System.Collections.Generic;
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
            
            var htmlContentElement = generateNewElement ? new XElement(srcHtmlContentElement) : srcHtmlContentElement ;
            htmlContentElement.Element(WordMl.SdtPrName)
                .Element(WordMl.TagName)
                .Attribute(WordMl.ValAttributeName)
                .SetValue(ProcessedHtmlContentTagName);
            htmlContent = HttpUtility.HtmlDecode(htmlContent);
            htmlContentElement.Element(WordMl.SdtContentName).Value = htmlContent;            
            return htmlContentElement;
        }
        
        public override void Process()
        {            
            base.Process();
            var element = this.HtmlTag.TagNode;
            MakeHtmlContentProcessed(element, DataReader.ReadText(this.HtmlTag.Expression));

        }
    }
}
