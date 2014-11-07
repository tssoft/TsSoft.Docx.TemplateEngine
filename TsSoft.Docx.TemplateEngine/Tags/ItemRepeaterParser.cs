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

        public bool IsBeforeNestedRepeater
        {
            get { return this.NextNestedRepeater != null; }
        }

        public bool HasExpression
        {
            get { return !string.IsNullOrEmpty(this.Expression); }
        }

        public bool HasElements
        {
            get { return this.Elements != null && this.Elements.Any(); }
        }

        public bool IsItem { get; set; }

        public bool IsItemIf { get; set; }

        public bool IsItemRepeater { get; set; }

        public bool IsEndItemRepeater { get; set; }

    }

    internal class ItemRepeaterTag
    {
        public XElement StartItemRepeater { get; set; }

        public XElement EndItemRepeater { get; set; }

        public bool IsNotSeparatedRepeater { get; set; }

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
                        IsEndItemRepeater = element.IsTag("enditemrepeater"),
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
            
            var nestedRepeaters = this.MarkNotSeparatedRepeaters(this.GetAllNestedRepeaters(tag), itemRepeaterElements); 
            if (startElement.Parent.Name == WordMl.ParagraphName) //&& ((itemRepeaterElements.Count != 0) && itemRepeaterElements.All(nr => nr.XElement.Parent.Name != WordMl.ParagraphName)))
            {
                startElement = startElement.Parent;
            }                                   
            XElement current = startElement;
            var repeaterElements = this.GetSiblingElements(itemRepeaterElements.ToList(), nestedRepeaters.ToList()).Where(sel => !(sel.XElement.IsTag("enditemrepeater") || sel.XElement.IsTag("itemrepeater")));            
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
                foreach (
                    var itemRepeaterElement in
                        itemRepeaterElements)
                {                   
                    var flagAdd = false;
                    foreach (var nestedRepeater in nestedRepeaters)
                    {
                        if (
                            !TraverseUtils.ElementsBetween(nestedRepeater.StartItemRepeater,
                                                           (nestedRepeater.EndItemRepeater.Parent.Name ==
                                                            WordMl.ParagraphName)
                                                               ? nestedRepeater.EndItemRepeater.Parent
                                                               : nestedRepeater.EndItemRepeater)
                                          .Any(element => element.Equals(itemRepeaterElement.XElement)))
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
                        if (itemRepeaterElement.HasElements && (itemRepeaterElement.Elements.Any(ire => ire.IsItemRepeater) || itemRepeaterElement.Elements.Any(ire => ire.IsEndItemRepeater)) && itemRepeaterElement.XElement.Name.Equals(WordMl.ParagraphName))
                        {
                            repeaterElements.AddRange(
                                itemRepeaterElement.Elements.Where(
                                    ire =>
                                    /*ire.XElement.IsSdt() &&*/ !ire.XElement.IsTag("itemrepeater") &&
                                    !ire.XElement.IsTag("enditemrepeater")));
                        }
                        else
                        {
                            repeaterElements.Add(itemRepeaterElement);
                        }
                    }               
            }                
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
                if ((nextRepeaterElement != null) && nextRepeaterElement.XElement.Name.Equals(WordMl.ParagraphName) && nextRepeaterElement.HasElements && nextRepeaterElement.Elements.First(ire => !ire.XElement.Name.Equals(WordMl.ParagraphPropertiesName)).IsItemRepeater)
                {
                    repeaterElement.NextNestedRepeater =
                        nextRepeaterElement.Elements.First(
                            ire => !ire.XElement.Name.Equals(WordMl.ParagraphPropertiesName));
                }

                if (repeaterElement.HasElements)
                {
                    repeaterElement.Elements = this.MarkLastElements(repeaterElement.Elements);
                }
                /*if (repeaterElement.XElement.DescendantsAndSelf().Any(el => el.IsTag("enditemrepeater")) && repeaterElement.IsBeforeNestedRepeater)
                {
                    repeaterElements.ElementAt(i - 1).NextNestedRepeater = repeaterElement.NextNestedRepeater;
                }*/
            }
            return repeaterElements;
        }

        private IEnumerable<ItemRepeaterTag> MarkNotSeparatedRepeaters(IEnumerable<ItemRepeaterTag> tags,
                                                                       IEnumerable<ItemRepeaterElement> repeaterElements)
        {
            repeaterElements = repeaterElements.ToList();
            tags = tags.ToList();
            for (var i = 0; i < repeaterElements.Count(); i++)
            {
                var repeaterElement = repeaterElements.ElementAt(i);
                var nextRepeaterElement = (i + 1 == repeaterElements.Count()) ? null : repeaterElements.ElementAt(i + 1);
                if ((repeaterElement.IsEndItemRepeater || repeaterElement.Elements.Any(re => re.IsEndItemRepeater)) && (nextRepeaterElement != null) &&
                    repeaterElement.IsBeforeNestedRepeater)
                {
                    tags.SingleOrDefault(t => t.StartItemRepeater.Equals(nextRepeaterElement.XElement)).IsNotSeparatedRepeater = true;                    
                }
            }            
            return tags;
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
                                                         (tag.EndItemRepeater.Parent.Name == WordMl.ParagraphName) ? tag.EndItemRepeater.Parent : tag.EndItemRepeater).Select(MakeElementCallback)                                        
                                        .ToList();
            var dataReaders = dataReader.GetReaders(tag.Source).ToList();
            for (var index = 1; index <= dataReaders.Count; index++)
            {
                ICollection<XElement> bisectElements;
                current = this.ProcessElements(elements, dataReaders[index - 1], current, null, index, out bisectElements);
            }
            return current;
                                  
        }

        private XElement ProcessNestedRepeaters(ItemRepeaterTag tag, DataReader dataReader, ICollection<XElement> bisectElements)
        {
            XElement current = null;
            int i = 0;
            foreach (var nestedRepeater in tag.NestedRepeaters)
            {
                if (nestedRepeater.NestedRepeaters.Count == 0)
                {
                    if (!nestedRepeater.IsNotSeparatedRepeater)
                    {
                        current = bisectElements.ElementAt(i);
                        ++i;
                    }                    
                    current = this.RenderDataReaders(nestedRepeater, dataReader, current);
                }
                else
                {
                    /*
                    this.ProcessNestedRepeaters(nestedRepeater, dataReader, currentElement);
                    current = this.RenderDataReaders(nestedRepeater, dataReader, currentElement);
                     */
                }
                
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
                        ICollection<XElement> bisectNestedElements = new List<XElement>();
                        this.ProcessElements(itemRepeaterElement.Elements, dataReader, null, result, index, out bisectNestedElements, true);
                        if (bisectNestedElements.Count != 0)
                        {
                            tempElementsBeforeItemRepeaters.Add(bisectNestedElements.ElementAt(0));
                        }
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
