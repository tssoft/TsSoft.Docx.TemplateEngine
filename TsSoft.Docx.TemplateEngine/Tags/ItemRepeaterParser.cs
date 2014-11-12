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
            if (startElement.Parent.Name == WordMl.ParagraphName)
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
                        if (itemRepeaterElement.HasElements 
                            && (itemRepeaterElement.Elements.Any(ire => ire.IsItemRepeater)  || itemRepeaterElement.Elements.Any(ire => ire.IsEndItemRepeater)) 
                            && itemRepeaterElement.XElement.Name.Equals(WordMl.ParagraphName))
                        {
                            var nestedElements = itemRepeaterElement.Elements.ToList();
                            var element = new XElement(itemRepeaterElement.XElement.Name);
                            var dicBisect = new Dictionary<XElement, ItemRepeaterElement>();
                            ItemRepeaterElement val;
                            if (!nestedElements.Any(ne => ne.IsEndItemRepeater))
                            {
                                foreach (var nestedElement in nestedElements)
                                {
                                    element.Add(nestedElement.XElement);
                                    if (nestedElement.IsBeforeNestedRepeater)
                                    {
                                        dicBisect.Add(nestedElement.XElement, new ItemRepeaterElement() { XElement = nestedElement.NextNestedRepeater.XElement });
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                var end = nestedElements.First(ne => ne.IsEndItemRepeater);
                                var next = end.XElement;
                                while ((next = next.NextElement()) != null)
                                {
                                    element.Add(next);
                                }
                            }
                            var newItemRepeaterElement = MakeElementCallback(element);
                            var elements = newItemRepeaterElement.Elements.ToList();

                            foreach (var kv in dicBisect)
                            {
                                foreach (var repeaterElement in elements)
                                {
                                    if (XNode.DeepEquals(repeaterElement.XElement, kv.Key))
                                    {
                                        val = new ItemRepeaterElement { XElement = new XElement(kv.Value.XElement) };
                                        repeaterElement.NextNestedRepeater = val;
                                    }
                                }
                            }
                            newItemRepeaterElement.Elements = elements;
                            
                            if (element.HasElements)
                            {
                                repeaterElements.Add(newItemRepeaterElement);
                            }
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
            }
            return repeaterElements;
        }

        private IEnumerable<ItemRepeaterTag> MarkNotSeparatedRepeaters(
                                                                       IEnumerable<ItemRepeaterTag> tags,
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
                    ((tag.EndItemRepeater.Parent.Name == WordMl.ParagraphName) 
                    && (!tag.EndItemRepeater.Parent.Elements().Contains(tag.StartItemRepeater))
                    && (tag.StartItemRepeater.Parent.Elements().Any(el => el.IsSdt() && !el.IsTag("itemrepeater")))
                    ) ? tag.EndItemRepeater.Parent : tag.EndItemRepeater)
                    .Select(MakeElementCallback)                                        
                    .ToList();
            IList<ItemRepeaterElement> processedElements = new List<ItemRepeaterElement>();            
            if (elements.Count > 1)
            {
                XElement tmpContainer = null;
                var tmpParent = elements[0].XElement.Parent; 
                IList<XElement> tmpElements = new List<XElement>();
                if (tmpParent.Name.Equals(WordMl.ParagraphName))
                {
                    tmpElements.Add(elements[0].XElement);
                }
                else
                {
                    processedElements.Add(elements[0]);
                }
                for (var i = 1; i < elements.Count(); i++)
                {
                    if (!elements[i].XElement.Parent.Name.Equals(WordMl.ParagraphName))
                    {
                        processedElements.Add(elements[i]);
                        continue;                        
                    }
                    if (elements[i].XElement.Parent.Equals(tmpParent))
                    {
                        tmpElements.Add(elements[i].XElement);
                    }
                    if (!elements[i].XElement.Parent.Equals(tmpParent) || (i == elements.Count - 1))
                    {
                        tmpContainer = new XElement(tmpParent.Name);
                        tmpContainer.Add(tmpElements);
                        processedElements.Add(MakeElementCallback(tmpContainer));
                        tmpParent = elements[i].XElement.Parent;   
                        tmpElements.Clear();              
                    }                    
                }
            }
            else
            {
                processedElements.Add(elements[0]);
            }
            
            var flgWrapInline = false;
            string inlineSeparator = string.Empty;
            if (processedElements.Count == 1)
            {
                var rootContainer = processedElements.FirstOrDefault();
                if ((rootContainer != null) && rootContainer.XElement.Name.Equals(WordMl.ParagraphName))
                {
                    var originalContainer = elements[0].XElement.Parent;
                    var endItemRepeater = originalContainer.Elements().FirstOrDefault(el => el.IsTag("enditemrepeater"));
                    if (endItemRepeater != null)
                    {
                        foreach (var itemRepeaterElement in elements)
                        {
                            if (itemRepeaterElement.XElement.Parent.Elements().Contains(endItemRepeater))
                            {
                                flgWrapInline = true;
                            }
                            else
                            {
                                flgWrapInline = false;
                                break;                                
                            }
                        }
                        if (flgWrapInline)
                        {
                            inlineSeparator = endItemRepeater.GetExpression();
                        }
                    }                    
                }
            }
            var dataReaders = dataReader.GetReaders(tag.Source).ToList();
            for (var index = 1; index <= dataReaders.Count; index++)
            {
                ICollection<XElement> bisectElements;
                var inlineWrapping = (index > 1) && flgWrapInline;
                if (inlineWrapping)
                {
                    current = current.Elements().Last();
                    var separatorElement = this.CreateSeparatorElement(inlineSeparator);
                    current.AddAfterSelf(separatorElement);
                    current = separatorElement;
                }
                current = this.ProcessElements(
                    inlineWrapping ? processedElements.First().Elements.ToList() : processedElements,
                    dataReaders[index - 1],
                    inlineWrapping ? null : current, 
                    inlineWrapping ? current.Parent : null,
                    index,
                    out bisectElements,
                    (index > 1) && flgWrapInline);                                
            }
            return current;
                                  
        }

        private XElement CreateSeparatorElement(string separator)
        {
            const string SpecialChar = "\'";            
            separator = separator.Replace(SpecialChar, "'");
            int quotePos = -1;
            while ((quotePos = separator.LastIndexOf('\'')) > -1)
            {
                separator = separator.Remove(quotePos, 1);
            }            
            var result = new XElement(WordMl.TextRunName, new XElement(WordMl.TextName, separator));
            return result;
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
                                                          !nested
                        );
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
