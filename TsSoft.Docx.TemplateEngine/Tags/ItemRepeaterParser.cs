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
        public XElement StartItemRepeater { get; set; }

        public XElement EndItemRepeater { get; set; }

        public string Source { get; set; }

        public ICollection<ItemRepeaterTag> NestedRepeaters { get; set; }

        public Func<XElement, RepeaterElement> MakeElementCallback { get; set; }
    }

    internal class ItemRepeaterParser
    {
        private const string IndexTag = "RItemIndex";
        private const string ItemTag = "RItemText";
        private const string ItemIf = "RItemIf";

        private static Func<XElement, ItemRepeaterElement> MakeElementCallback = element =>
            {
               // if (element.IsSdt() && element.IsTag("itemrepeater"))
                //{
                  //  return null;
                //}
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
        

        public void Parse(ItemRepeaterTag tag, IList<DataReader> dataReaders)
        {           
            var startElement = tag.StartItemRepeater;
            var endElement = tag.EndItemRepeater;
            var itemRepeaterElements =
                TraverseUtils.ElementsBetween(startElement, endElement).Select(MakeElementCallback).ToList();

            XElement current = startElement;
            for (var index = 1; index <= dataReaders.Count(); index++)
            {
                foreach (var nestedRepeater in tag.NestedRepeaters)
                {
                    this.Parse(nestedRepeater, dataReaders[index - 1].GetReaders(nestedRepeater.Source).ToList());
                }
                var repeaterElements = itemRepeaterElements as IList<ItemRepeaterElement> ??
                                        itemRepeaterElements.ToList();
                current = this.ProcessElements(repeaterElements, dataReaders[index - 1], current, null, index);
            }
            foreach (var itemRepeaterElement in itemRepeaterElements)
            {
                itemRepeaterElement.XElement.Remove();
            }
            startElement.Remove();
            endElement.Remove();
            
        }

        private void ProcessNestedRepeaters(ItemRepeaterTag tag)
        {
            foreach (var nestedRepeater in tag.NestedRepeaters)
            {
                if (nestedRepeater.NestedRepeaters.Count == 0)
                {
                    var dataReaders = new DataReader()
                    for (int index = 1; index )
                }
            }
        }

        private void RemoveTags()
        {
            
        }

        private XElement ProcessElements(IEnumerable<ItemRepeaterElement> itemRepeaterElements, DataReader dataReader, XElement start, XElement parent, int index)
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
                        this.ProcessElements(itemRepeaterElement.Elements, dataReader, null, result, index);
                    }
                    else
                    {
                        element.Value = itemRepeaterElement.XElement.Value;
                    }
                }
                if (start != null)
                {
                    start.AddAfterSelf(result);
                    start = result;
                }
                else
                {
                    parent.Add(result);
                }
            }
            return result;            
        }
    }
}
