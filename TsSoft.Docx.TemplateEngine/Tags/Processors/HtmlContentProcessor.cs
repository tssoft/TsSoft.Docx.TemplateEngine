using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    internal class HtmlContentProcessor : AbstractProcessor
    {
        public static string ProcessedHtmlContentTagName = "processedhtmlcontent";

        public HtmlContentTag HtmlTag { get; set; }
        
        public override void Process()
        {            
            base.Process();
            var element = this.HtmlTag.TagNode;
            element.Element(WordMl.SdtPrName).Element(WordMl.TagName).Attribute(WordMl.ValAttributeName).SetValue(ProcessedHtmlContentTagName);
            var htmlContent = HttpUtility.HtmlDecode(DataReader.ReadText(this.HtmlTag.Expression));            
            element.Element(WordMl.SdtContentName).Value = htmlContent;
        }
    }
}
