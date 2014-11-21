using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
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
        private const string EndItemIf = "REndIf";
        private const string ItemHtmlContentTag = "RItemHtmlContent";
        private static Func<XElement, ItemRepeaterElement> MakeElementCallback = element =>
            {
                var itemRepeaterElement = new ItemRepeaterElement()
                    {
                        Elements = element.Elements().Select(MakeElementCallback),
                        IsIndex = element.IsTag(IndexTag),
                        IsItem = element.IsTag(ItemTag),
                        IsItemIf = element.IsTag(ItemIf),
                        IsEndItemIf = element.IsTag(EndItemIf),
                        IsItemRepeater = element.IsTag("itemrepeater"),
                        IsEndItemRepeater = element.IsTag("enditemrepeater"),
                        IsItemHtmlContent = element.IsTag(ItemHtmlContentTag),
                        XElement = element
                    };
                if (itemRepeaterElement.IsItem || itemRepeaterElement.IsItemIf)
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
                this.SetRepeaterElementsVisible(repeaterElements.ToList());
                ICollection<XElement> bisectElements;
                var ifCount = 1;
                var flg = false;
                repeaterElements = this.MarkInvisibleElements(repeaterElements, dataReaders[index - 1], ref ifCount, ref flg, index == dataReaders.Count);
                if (ifCount != 1)
                {
                    throw new Exception("ItemIf error. Check REndIf count.");
                }
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
                foreach (var itemRepeaterElement in itemRepeaterElements)
                {
                    var flagAdd = nestedRepeaters.All(nr =>
                        {
                            var elementsBetween = TraverseUtils.ElementsBetween(nr.StartItemRepeater,
                                                                                nr.EndItemRepeater.Parent != null && (nr.EndItemRepeater.Parent.Name == WordMl.ParagraphName)
                                                                                    ? nr.EndItemRepeater.Parent : nr.EndItemRepeater);
                            return !elementsBetween.Any(element => element.Equals(itemRepeaterElement.XElement));
                        });
                    if (flagAdd)
                    {
                        if (itemRepeaterElement.HasElements
                            && (itemRepeaterElement.Elements.Any(ire => ire.IsItemRepeater) || itemRepeaterElement.Elements.Any(ire => ire.IsEndItemRepeater))
                            && itemRepeaterElement.XElement.Name.Equals(WordMl.ParagraphName))
                        {
                            var nestedElements = itemRepeaterElement.Elements.ToList();
                            var element = new XElement(itemRepeaterElement.XElement.Name);
                            var dicBisect = new Dictionary<XElement, ItemRepeaterElement>();
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
                                foreach (var repeaterElement in elements.Where(repeaterElement => XNode.DeepEquals(repeaterElement.XElement, kv.Key)))
                                {
                                    var val = new ItemRepeaterElement { XElement = new XElement(kv.Value.XElement) };
                                    repeaterElement.NextNestedRepeater = val;
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

        private IEnumerable<ItemRepeaterElement> SetRepeaterElementsVisible(IEnumerable<ItemRepeaterElement> repeaterElements)
        {
            var resultElements = new List<ItemRepeaterElement>(repeaterElements);
            foreach (var itemRepeaterElement in resultElements)
            {
                itemRepeaterElement.IsVisible = true;
                if (itemRepeaterElement.HasElements)
                {
                    itemRepeaterElement.Elements = this.SetRepeaterElementsVisible(itemRepeaterElement.Elements.ToList());
                }
            }
            return resultElements;
        }

        private IEnumerable<ItemRepeaterElement> MarkInvisibleElements(IEnumerable<ItemRepeaterElement> repeaterElements, DataReader dataReader, ref int ifCount, ref bool flgProcessingInvisibleContent, bool isLast)
        {
            const string IsLastConditionName = "isnotlastelement";
            var resultRepeaterElements = new List<ItemRepeaterElement>(repeaterElements);
            foreach (var itemRepeaterElement in resultRepeaterElements)
            {
                if (flgProcessingInvisibleContent && (!itemRepeaterElement.XElement.Name.Equals(WordMl.ParagraphName) || itemRepeaterElement.XElement.IsSdt()))
                {
                    itemRepeaterElement.IsVisible = false;
                }
                if (itemRepeaterElement.IsItemIf)
                {
                    if (itemRepeaterElement.IsVisible)
                    {
                        var expression = itemRepeaterElement.Expression.ToLower();
                        var condition = expression.Equals(IsLastConditionName) ? !isLast : bool.Parse(dataReader.ReadText(expression));
                        flgProcessingInvisibleContent = !condition;
                    }
                    else
                    {
                        ifCount++;
                    }
                }
                else if (itemRepeaterElement.IsEndItemIf)
                {
                    if (ifCount == 1)
                    {
                        flgProcessingInvisibleContent = false;
                    }
                    else
                    {
                        ifCount--;
                    }
                }
                else
                {
                    if (itemRepeaterElement.HasElements)
                    {
                        itemRepeaterElement.Elements = this.MarkInvisibleElements(itemRepeaterElement.Elements, dataReader, ref ifCount, ref flgProcessingInvisibleContent, isLast);
                    }
                }
            }
            return resultRepeaterElements;
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

        private IEnumerable<ItemRepeaterElement> CreateParagraphForInlineElements(IList<ItemRepeaterElement> elements)
        {
            var processedElements = new List<ItemRepeaterElement>();
            elements = elements.ToList();
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
                var parent = elements[i].XElement.Parent;
                if (!parent.Name.Equals(WordMl.ParagraphName))
                {
                    processedElements.Add(elements[i]);
                    continue;
                }
                if (parent.Equals(tmpParent))
                {
                    tmpElements.Add(elements[i].XElement);
                }
                if (!parent.Equals(tmpParent) || (i == elements.Count - 1))
                {
                    XElement tmpContainer = new XElement(tmpParent.Name);
                    tmpContainer.Add(tmpElements);
                    processedElements.Add(MakeElementCallback(tmpContainer));
                    tmpParent = parent;
                    tmpElements.Clear();
                }
            }
            return processedElements;
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
            var processedElements = new List<ItemRepeaterElement>();
            if (elements.Count > 1)
            {
                processedElements.AddRange(this.CreateParagraphForInlineElements(elements));
            }
            else
            {
                processedElements.Add(elements[0]);
            }
            processedElements = this.SetRepeaterElementsVisible(processedElements.ToList()).ToList();
            var flgWrapInline = false;
            if (processedElements.Count == 1)
            {
                var rootContainer = processedElements.FirstOrDefault();
                if ((rootContainer != null) && rootContainer.XElement.Name.Equals(WordMl.ParagraphName))
                {
                    var originalContainer = elements[0].XElement.Parent;
                    var endItemRepeater = originalContainer.Elements().FirstOrDefault(el => el.IsTag("enditemrepeater"));
                    if (endItemRepeater != null)
                    {
                        flgWrapInline = elements.All(e => e.XElement.Parent.Elements().Contains(endItemRepeater));
                    }
                }
            }
            var dataReaders = dataReader.GetReaders(tag.Source).ToList();
            for (var index = 1; index <= dataReaders.Count; index++)
            {
                ICollection<XElement> bisectElements;
                var ifCount = 1;
                var flg = false;
                processedElements = this.MarkInvisibleElements(processedElements, dataReaders[index - 1], ref ifCount, ref flg, index == dataReaders.Count).ToList();
                if (ifCount != 1)
                {
                    throw new Exception("RItemIf error. Check REndIf count.");
                }
                var inlineWrapping = (index > 1) && flgWrapInline;
                if (inlineWrapping)
                {
                    current = current.Elements().Last();
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
                if (!itemRepeaterElement.IsVisible || itemRepeaterElement.IsItemIf || itemRepeaterElement.IsEndItemIf)
                {
                    continue;
                }                
                if (itemRepeaterElement.IsItemHtmlContent)
                {
                    result = HtmlContentProcessor.MakeHtmlContentProcessed(itemRepeaterElement.XElement, dataReader.ReadText(itemRepeaterElement.Expression), true);
                }
                else if (itemRepeaterElement.IsIndex)
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
