using System;
using System.Collections;
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

        public ItemRepeaterElement NextNestedRepeater { get; set; }

        public bool IsBeforeNestedRepeater {
            get { return this.NextNestedRepeater != null; }
        }

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

        public bool IsItemRepeater { get; set; }

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
                        IsItemRepeater = element.IsTag("itemrepeater"),
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
                this.MarkLastElements(TraverseUtils.ElementsBetween(startElement, endElement).Select(MakeElementCallback).ToList()).ToList();
            
            var nestedRepeaters = this.GetAllNestedRepeaters(tag); 
            if (startElement.Parent.Name == WordMl.ParagraphName) //&& ((itemRepeaterElements.Count != 0) && itemRepeaterElements.All(nr => nr.XElement.Parent.Name != WordMl.ParagraphName)))
            {
                startElement = startElement.Parent;
            }                                   
            XElement current = startElement;
            var repeaterElements = this.GetSiblingElements(itemRepeaterElements.ToList(), nestedRepeaters.ToList());            
            XElement currentNested;       
            for (var index = 1; index <= dataReaders.Count(); index++)
            {
                ICollection<XElement> bisectElements;
                current = this.ProcessElements(repeaterElements, dataReaders[index - 1], current, null, index, out bisectElements);                
                currentNested = this.ProcessNestedRepeaters(tag, dataReaders[index - 1], (!bisectElements.Any()) ? new List<XElement>() { current } : bisectElements);
                if ((currentNested != null) && repeaterElements.Last().IsBeforeNestedRepeater)
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

        private IEnumerable<ItemRepeaterElement> GetSiblingElements(
            IEnumerable<ItemRepeaterElement> itemRepeaterElements, IEnumerable<ItemRepeaterTag> nestedRepeaters)
        {            
            var repeaterElements = new List<ItemRepeaterElement>();
            int i = 0;
            if (!nestedRepeaters.Any())
            {
                repeaterElements.AddRange(itemRepeaterElements);
            }
            else
            {
                foreach (var itemRepeaterElement in itemRepeaterElements.Where(ire => !(ire.XElement.IsTag("itemrepeater") || ire.XElement.IsTag("enditemrepeater"))))
                {                    
                    var flagAdd = false;
                    foreach (var nestedRepeater in nestedRepeaters)
                    {
                        if (
                            !TraverseUtils.ElementsBetween(nestedRepeater.StartItemRepeater,
                                                           nestedRepeater.EndItemRepeater)
                                          .Contains(itemRepeaterElement.XElement))
                        {
                            flagAdd = true;
                        }
                        else
                        {
                            flagAdd = false;
                            break;                            
                        }
                    }
                    if (flagAdd)
                    {
                        repeaterElements.Add(itemRepeaterElement);
                    }                    
                }
                i++;
            }
            return repeaterElements;
        }
       
        private IEnumerable<ItemRepeaterElement> MarkLastElements(IEnumerable<ItemRepeaterElement> repeaterElements)
        {
            repeaterElements = repeaterElements.ToList();
            for (var i = 0; i < repeaterElements.Count(); i++)
            {
                var repeaterElement = repeaterElements.ElementAt(i);
                var nextRepeaterElement = (i + 1 == repeaterElements.Count()) ? null : repeaterElements.ElementAt(i + 1);
                repeaterElement.NextNestedRepeater = ((nextRepeaterElement != null) &&
                                                      nextRepeaterElement.IsItemRepeater)
                                                         ? nextRepeaterElement
                                                         : null;
            }
            return repeaterElements;
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

        private XElement RenderDataReaders(ItemRepeaterTag tag, DataReader dataReader, XElement current)
        {            
            var elements = TraverseUtils.ElementsBetween(tag.StartItemRepeater,
                                                         tag.EndItemRepeater).Select(MakeElementCallback)                                        
                                        .ToList();
            var dataReaders = dataReader.GetReaders(tag.Source).ToList();
            for (var index = 1; index <= dataReaders.Count; index++)
            {
                ICollection<XElement> bisectElements;
                current = this.ProcessElements(elements, dataReaders[index - 1], current, null, index, out bisectElements);
            }
            return current;
                                  
        }

        private XElement ProcessNestedRepeaters(ItemRepeaterTag tag, DataReader dataReader, ICollection<XElement> currentElements)
        {
            XElement current = null;
            int i = 0;
            foreach (var nestedRepeater in tag.NestedRepeaters)
            {
                if (nestedRepeater.NestedRepeaters.Count == 0)
                {
                    current = currentElements.ElementAt(i);
                    current = this.RenderDataReaders(nestedRepeater, dataReader, current);
                }
                else
                {
                    /*
                    this.ProcessNestedRepeaters(nestedRepeater, dataReader, currentElement);
                    current = this.RenderDataReaders(nestedRepeater, dataReader, currentElement);
                     */
                }
                ++i;
            }
            return current;
        }

        private void RemoveTags()
        {
            
        }

        private XElement ProcessElements(IEnumerable<ItemRepeaterElement> itemRepeaterElements, DataReader dataReader, XElement start, XElement parent, int index, out ICollection<XElement> elementsBeforeNestedRepeaters, bool nested = false)
        {
            XElement result = null;
            XElement previous = start;
            ICollection<XElement> tempElementsBeforeItemRepeaters = new List<XElement>();
            foreach (var itemRepeaterElement in itemRepeaterElements)
            {                
                if (itemRepeaterElement.IsIndex)
                {
                    result = DocxHelper.CreateTextElement(itemRepeaterElement.XElement,
                                                          itemRepeaterElement.XElement.Parent,
                                                          index.ToString(CultureInfo.CurrentCulture),
                                                          !nested);
                }
                else if (itemRepeaterElement.IsItem && itemRepeaterElement.HasExpression)
                {
                    result = DocxHelper.CreateTextElement(itemRepeaterElement.XElement,
                                                          itemRepeaterElement.XElement.Parent,
                                                          dataReader.ReadText(itemRepeaterElement.Expression),
                                                          !nested
                        );
                }
                else
                {                    
                    var element = new XElement(itemRepeaterElement.XElement);
                    element.RemoveNodes();
                    result = element;
                    if (itemRepeaterElement.HasElements)
                    {
                        this.ProcessElements(itemRepeaterElement.Elements, dataReader, null, result, index, out tempElementsBeforeItemRepeaters, true);
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
                if (itemRepeaterElement.IsBeforeNestedRepeater)
                {
                    tempElementsBeforeItemRepeaters.Add(result);                    
                }
                
            }
            elementsBeforeNestedRepeaters = tempElementsBeforeItemRepeaters;
            return result;            
        }
    }
}
