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
            var nestedRepeaters = this.GetAllNestedRepeaters(tag);
            XElement current = startElement;
            var repeaterElements = new List<ItemRepeaterElement>();
            if (nestedRepeaters.Count == 0)
            {
                repeaterElements.AddRange(itemRepeaterElements);
            }
            else
            {
                repeaterElements.AddRange(from itemRepeaterElement in itemRepeaterElements.Where(ire => !(ire.XElement.IsTag("itemrepeater") || ire.XElement.IsTag("enditemrepeater")))
                                          from nestedRepeater in nestedRepeaters
                                          where (!TraverseUtils.ElementsBetween(nestedRepeater.StartItemRepeater, nestedRepeater.EndItemRepeater).Contains(itemRepeaterElement.XElement))
                                          select itemRepeaterElement);
            }
            XElement currentNested;
            for (var index = 1; index <= dataReaders.Count(); index++)
            {
                /*if (index == 1)
                {
                    current = this.ProcessNestedRepeaters(tag, dataReaders[index - 1]);
                }
                else
                {
                    this.ProcessNestedRepeaters(tag, dataReaders[index - 1]);
                }*/
                current = this.ProcessElements(repeaterElements, dataReaders[index - 1], current, null, index);
                currentNested = this.ProcessNestedRepeaters(tag, dataReaders[index - 1], current);
                if (currentNested != null)
                {
                    current = currentNested;
                }

            }            
            foreach (var itemRepeaterElement in itemRepeaterElements)
            {
                itemRepeaterElement.XElement.Remove();
            }
            startElement.Remove();
            endElement.Remove();
            
        }

        private ICollection<ItemRepeaterTag> GetAllNestedRepeaters(ItemRepeaterTag tag)
        {
            var result = new List<ItemRepeaterTag>();
            foreach (var nestedRepeater in tag.NestedRepeaters)
            {
                if (nestedRepeater.NestedRepeaters.Count != 0)
                {
                    result.AddRange(this.GetAllNestedRepeaters(nestedRepeater));
                }
                else
                {
                    result.Add(nestedRepeater);
                }
            }
            return result;
        }

        private XElement RenderDataReaders(ItemRepeaterTag tag, DataReader dataReader,XElement current)
        {
            var elements = TraverseUtils.ElementsBetween(tag.StartItemRepeater,
                                                         tag.EndItemRepeater).Select(MakeElementCallback)
                                        .Where(
                                            el =>
                                            el.XElement.IsSdt() || el.XElement.Descendants()
                                                                    .Any(nel => nel.IsSdt()))
                                        .ToList();
            var dataReaders = dataReader.GetReaders(tag.Source).ToList();
            //var current = tag.StartItemRepeater;
            for (var index = 1; index <= dataReaders.Count; index++)
            {
                current = this.ProcessElements(elements, dataReaders[index - 1], current, null, index);
            }
            return current;
                                  
        }

        private XElement ProcessNestedRepeaters(ItemRepeaterTag tag, DataReader dataReader, XElement curr)
        {
            XElement current = null;
            foreach ( var nestedRepeater in tag.NestedRepeaters)
            {
                if (nestedRepeater.NestedRepeaters.Count == 0)
                {                    
                    current = this.RenderDataReaders(nestedRepeater, dataReader, curr);
                    //current = nestedRepeater.EndItemRepeater;
                }
                else
                {
                    this.ProcessNestedRepeaters(nestedRepeater, dataReader, curr);
                    current = this.RenderDataReaders(nestedRepeater, dataReader, curr);
                }                
            }
            return current;
        }

        private void RemoveTags()
        {
            
        }

        private XElement ProcessElements(IEnumerable<ItemRepeaterElement> itemRepeaterElements, DataReader dataReader, XElement start, XElement parent, int index)
        {
            XElement result = null;
            XElement previous = start;            
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
    }
}
