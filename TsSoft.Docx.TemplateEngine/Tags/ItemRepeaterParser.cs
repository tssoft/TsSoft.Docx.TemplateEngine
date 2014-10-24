using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;
namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal class ItemRepeaterElement 
    {
        public IEnumerable<ItemRepeaterElement> Elements { get; set; }

        public XElement XElement { get; set; }

        public string Expression { get; set; }

        public bool IsIndex { get; set; }

        public bool HasExpression
        {
            get
            {
                return !string.IsNullOrEmpty(this.Expression);
            }
        }

        public bool HasElements
        {
            get
            {
                return this.Elements != null && this.Elements.Any();
            }
        }

        public bool IsItem { get; set; }

        public bool IsItemIf { get; set; }

    }

    internal class ItemRepeaterTag
    {
        public Func<XElement, RepeaterElement> MakeElementCallback { get; set; }
    }

    internal class ItemRepeaterParser
    {
        private const string IndexTag = "RItemIndex";
        private const string ItemTag = "RItemText";
        private const string ItemIf = "RItemIf";

        private static Func<XElement, ItemRepeaterElement> MakeElementCallback = element =>
            {
                var itemRepeaterElement = new ItemRepeaterElement()
                    {
                        Elements = element.Elements().Select(MakeElementCallback),
                        IsIndex = element.IsTag(IndexTag),
                        IsItem = element.IsTag(ItemTag),
                        IsItemIf = element.IsTag(ItemIf),
                        XElement = element
                    };
                if (itemRepeaterElement.IsItem)
                {
                    itemRepeaterElement.Expression = element.GetExpression();
                }
                return itemRepeaterElement;
            };

        public void Parse(XElement startElement, XElement endElement, IList<DataReader> dataReaders)
        {            
            var itemRepeaterElements = TraverseUtils.ElementsBetween(startElement, endElement).Select(MakeElementCallback);
           
            for (var index = 1; index < dataReaders.Count(); index++)
            {
                var repeaterElements = itemRepeaterElements as IList<ItemRepeaterElement> ?? itemRepeaterElements.ToList();
                this.ProcessElements(repeaterElements, dataReaders[index - 1], index);
            }

            foreach (var itemRepeaterElement in itemRepeaterElements)
            {
                itemRepeaterElement.XElement.Remove();
            }

            startElement.Remove();
            endElement.Remove();
            //throw new NotImplementedException();
        }        

        private void RemoveTags()
        {
            
        }

        private void ProcessElements(IEnumerable<ItemRepeaterElement> itemRepeaterElements, DataReader dataReader, int index)
        {
            XElement result = null;
            foreach (var itemRepeaterElement in itemRepeaterElements)
            {
                if (itemRepeaterElement.IsIndex)
                {
                    result = DocxHelper.CreateTextElement(itemRepeaterElement.XElement,
                                                          itemRepeaterElement.XElement.Parent,
                                                          index.ToString(CultureInfo.CurrentCulture));
                }
                else if (itemRepeaterElement.IsItem && itemRepeaterElement.HasExpression)
                {
                    result = DocxHelper.CreateTextElement(itemRepeaterElement.XElement,
                                                          itemRepeaterElement.XElement.Parent,
                                                          dataReader.ReadText(itemRepeaterElement.Expression));
                }
                else
                {
                    var element = new XElement(itemRepeaterElement.XElement);
                    element.RemoveNodes();
                    result = element;
                    if (itemRepeaterElement.HasElements)
                    {
                        this.ProcessElements(itemRepeaterElement.Elements, dataReader, index);
                    }
                    else
                    {
                        element.Value = itemRepeaterElement.XElement.Value;
                    }                    
                }                
                itemRepeaterElement.XElement.AddAfterSelf(result);
                itemRepeaterElement.XElement.Remove();
            }
        }
    }
}
