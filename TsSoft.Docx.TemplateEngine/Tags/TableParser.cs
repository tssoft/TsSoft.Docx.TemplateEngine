using System;
using System.Linq;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;

    /// <summary>
    /// Parse table tag
    /// </summary>
    internal class TableParser : GeneralParser
    {
        private static readonly string ItemIndexTagName = "itemindex";
        private static readonly string ItemTextTagName = "itemtext";
        private static readonly string ItemIfTagName = "itemif";
        private static readonly string EndItemIfTagName = "enditemif";
        private static readonly string ItemHtmlContentTagName = "itemhtmlcontent";
              
        /// <summary>
        /// Do parsing
        /// </summary>
        public override XElement Parse(ITagProcessor parentProcessor, XElement startElement)
        {
            this.ValidateStartTag(startElement, "Table");

            if (parentProcessor == null)
            {
                throw new ArgumentNullException();
            }

            var endTableTag = this.TryGetRequiredTags(startElement, "EndTable").First();
            var coreParser = new CoreTableParser(false);            
            var tag = coreParser.Parse(startElement, endTableTag);
            var processor = new TableProcessor { TableTag = tag };

            if (TraverseUtils.ElementsBetween(startElement, endTableTag).Any())
            {
                this.GoDeeper(processor, TraverseUtils.ElementsBetween(startElement, endTableTag).First());
            }
            parentProcessor.AddProcessor(processor);

            return endTableTag;
        }
                 
        private void GoDeeper(ITagProcessor parentProcessor, XElement element)
        {
            do
            {
                if (element.IsSdt())
                {
                    switch (this.GetTagName(element).ToLower())
                    {
                        case "itemtext":
                        case "itemindex":
                        case "itemif":
                        case "itemhtmlcontent":
                        case "itemrepeater":
                        case "enditemrepeater":
                        case "enditemif":
                            break;
                        default:
                            element = this.ParseSdt(parentProcessor, element);
                            break;
                    }
                }
                else if (element.HasElements)
                {
                    this.GoDeeper(parentProcessor, element.Elements().First());
                }
                element = element.NextElement();
            }
            while (element != null && (!element.IsSdt() || GetTagName(element).ToLower() != "endtable"));
        }
    }
}
