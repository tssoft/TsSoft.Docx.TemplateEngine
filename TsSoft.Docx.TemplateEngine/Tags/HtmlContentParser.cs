using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class HtmlContentParser : GeneralParser
    {
         public override XElement Parse(ITagProcessor parentProcessor, XElement startElement)
         {
             this.ValidateStartTag(startElement, "HtmlContent");
             var tag = new HtmlContentTag() { Expression = startElement.Value, TagNode = startElement };
             var processor = new HtmlContentProcessor() { HtmlTag = tag };
             parentProcessor.AddProcessor(processor);
             return startElement;
         }

    }    
}
