using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    internal class RepeaterProcessor : AbstractProcessor
    {
        public RepeaterTag RepeaterTag { get; set; }

        public override void Process()
        {
            base.Process();

            this.ProcessDynamicContent();

            var current = RepeaterTag.StartRepeater;
            var dataReaders = DataReader.GetReaders(RepeaterTag.Source).ToList();
            var repeaterElements =
                TraverseUtils.ElementsBetween(RepeaterTag.StartRepeater, RepeaterTag.EndRepeater)
                             .Select(RepeaterTag.MakeElementCallback).ToList();
            for (var index = 0; index < dataReaders.Count; index++)
            {
                current = this.ProcessElements(repeaterElements, dataReaders[index], current, null, index + 1);
            }
            foreach (var repeaterElement in repeaterElements)
            {
                repeaterElement.XElement.Remove();
            }

            if (this.LockDynamicContent)
            {
                var innerElements = TraverseUtils.ElementsBetween(this.RepeaterTag.StartRepeater, this.RepeaterTag.EndRepeater).ToList();
                innerElements.Remove();
                this.RepeaterTag.StartRepeater.AddBeforeSelf(DocxHelper.CreateDynamicContentElement(innerElements, this.RepeaterTag.StartRepeater));
                this.CleanUp(this.RepeaterTag.StartRepeater, this.RepeaterTag.EndRepeater);
            }
            else
            {
                this.RepeaterTag.StartRepeater.Remove();
                this.RepeaterTag.EndRepeater.Remove();
            }
        }

        private XElement ProcessElements(IEnumerable<RepeaterElement> elements, DataReader dataReader, XElement start, XElement parent, int index)
        {
            XElement result = null;
            XElement previous = start;


            foreach (var repeaterElement in elements)
            {
                if (repeaterElement.IsIndex)
                {
                    result = DocxHelper.CreateTextElement(repeaterElement.XElement, repeaterElement.XElement.Parent, index.ToString(CultureInfo.CurrentCulture));
                }
                else if (repeaterElement.IsItem && repeaterElement.HasExpression)
                {
                    result = DocxHelper.CreateTextElement(repeaterElement.XElement, repeaterElement.XElement.Parent, dataReader.ReadText(repeaterElement.Expression));
                }
                else
                {
                    var element = new XElement(repeaterElement.XElement);
                    element.RemoveNodes();
                    result = element;
                    if (repeaterElement.HasElements)
                    {
                        this.ProcessElements(repeaterElement.Elements, dataReader, null, result, index);
                    }
                    else
                    {
                        element.Value = repeaterElement.XElement.Value;
                    }
                }
                if (previous != null)
                {
                    previous.AddAfterSelf(result);                    
                    previous = result;
                }
                else
                {
                    parent.Add(result);
                }
            }
            return result;
        }

        private void ProcessDynamicContent()
        {
            var dynamicContentTags =
                TraverseUtils.ElementsBetween(this.RepeaterTag.StartRepeater, this.RepeaterTag.EndRepeater)
                             .Where(
                                 element =>
                                 element.IsSdt()
                                 && element.Element(WordMl.SdtPrName)
                                           .Element(WordMl.TagName)
                                           .Attribute(WordMl.ValAttributeName)
                                           .Value.ToLower()
                                           .Equals("dynamiccontent"))
                             .ToList();
            foreach (var dynamicContentTag in dynamicContentTags)
            {
                var innerElements = dynamicContentTag.Element(WordMl.SdtContentName).Elements();
                dynamicContentTag.AddAfterSelf(innerElements);
                dynamicContentTag.Remove();
            }
        }
    }
}
