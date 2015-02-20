using System;
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
                TraverseUtils.SecondElementsBetween(RepeaterTag.StartRepeater, RepeaterTag.EndRepeater)
                             .Select(RepeaterTag.MakeElementCallback).ToList();
            for (var index = 0; index < dataReaders.Count; index++)
            {
                XElement endIfElementTmp = null;
                current = this.ProcessElements(repeaterElements, dataReaders[index], current, null, index + 1, ref endIfElementTmp);
            }
            foreach (var repeaterElement in repeaterElements)
            {           
                repeaterElement.XElement.Remove();                
            }

            if (this.CreateDynamicContentTags)
            {
                var innerElements =
                    TraverseUtils.ElementsBetween(this.RepeaterTag.StartRepeater, this.RepeaterTag.EndRepeater).ToList();
                innerElements.Remove();
                this.RepeaterTag.StartRepeater.AddBeforeSelf(
                    DocxHelper.CreateDynamicContentElement(
                        innerElements, this.RepeaterTag.StartRepeater, this.DynamicContentLockingType));
                this.CleanUp(this.RepeaterTag.StartRepeater, this.RepeaterTag.EndRepeater);
            }
            else
            {
                this.RepeaterTag.StartRepeater.Remove();

                this.RepeaterTag.EndRepeater.Remove();
            }
        }
        private ItemRepeaterTag GenerateItemRepeaterTag(RepeaterElement itemRepeaterElement)
        {
            var tagResult = new ItemRepeaterTag()
            {
                StartItemRepeater = itemRepeaterElement.StartTag,
                EndItemRepeater = itemRepeaterElement.EndTag,
                NestedRepeaters = new List<ItemRepeaterTag>(),
                Source = itemRepeaterElement.Expression
            };
            foreach (var element in itemRepeaterElement.TagElements.Where(element => element.IsItemRepeater))
            {
                tagResult.NestedRepeaters.Add(this.GenerateItemRepeaterTag(element));
            }
            return tagResult;
        }

        private void ProcessItemRepeaterElement(RepeaterElement itemRepeaterElement, DataReader reader, int index,
                                                XElement previous)
        {
            var readers = reader.GetReaders(itemRepeaterElement.Expression);
            var itemRepeaterTag = GenerateItemRepeaterTag(itemRepeaterElement);
            var parser = new ItemRepeaterParser();
            parser.Parse(itemRepeaterTag, readers.ToList());            
        }

        private void ProcessItemIfElement(RepeaterElement ifElement, DataReader dataReader,
                                          ref XElement endIfElement)
        {
            var expression = ifElement.Expression;
            var condition = bool.Parse(dataReader.ReadText(expression));
            if (!condition)
            {
                endIfElement = ifElement.EndTag;
            }                                
        }

        private bool IsItemRepeaterElement(XElement element)
        {
            return element.Name.Equals(WordMl.ParagraphName)
                       ? element.Elements().All(el => ItemRepeaterGenerator.IsItemRepeaterElement(el))
                       : ItemRepeaterGenerator.IsItemRepeaterElement(element);
        }

        private XElement ProcessItemTableElement()
        {
            throw new NotImplementedException();
        }

        private XElement ProcessElements(IEnumerable<RepeaterElement> elements, DataReader dataReader, XElement start, XElement parent, int index, ref XElement endIfElement, bool nested = false)
        {
            XElement result = null;
            XElement previous = start;
            elements = elements.Where(el => !this.IsItemRepeaterElement(el.XElement)).ToList();            
            foreach (var repeaterElement in elements.Where(el => !this.IsItemRepeaterElement(el.XElement)).ToList())
            { 
                if (repeaterElement.IsEndItemIf && repeaterElement.Equals(endIfElement))
                {
                    endIfElement = null;
                    continue;
                }
                if ((endIfElement != null) && !repeaterElement.XElement.Name.Equals(WordMl.ParagraphName))
                {
                    continue;                    
                }
                if (repeaterElement.IsItemIf)
                {
                    this.ProcessItemIfElement(repeaterElement, dataReader, ref endIfElement);
                    continue;                    
                }                
                if (repeaterElement.IsEndItemTable || (repeaterElement.XElement.Name.Equals(WordMl.TableName) && repeaterElement.XElement.Descendants().Any(el => el.IsSdt())))
                {
                    continue;                    
                }
                if (repeaterElement.IsItemHtmlContent)
                {
                    result = HtmlContentProcessor.MakeHtmlContentProcessed(repeaterElement.XElement, dataReader.ReadText(repeaterElement.Expression), true);                    
                }
                else if (repeaterElement.IsItemTable)
                {
                    result = ItemTableGenerator.ProcessItemTableElement(repeaterElement.StartTag, repeaterElement.EndTag,
                                                                        dataReader);
                    if (nested)
                    {
                        previous.AddAfterSelf(result);
                        previous = result;
                        result = null;                        
                    }                    
                }
                else if (repeaterElement.IsItemRepeater)
                {                                        
                    var itemRepeaterTag = new ItemRepeaterTag()
                        {
                            StartItemRepeater = repeaterElement.StartTag,
                            EndItemRepeater = repeaterElement.EndTag,
                            Source = repeaterElement.Expression                            
                        };                                       
                    var itemRepeaterGenerator = new ItemRepeaterGenerator();
                    previous = itemRepeaterGenerator.Generate(itemRepeaterTag,
                                                              dataReader.GetReaders(repeaterElement.Expression),
                                                              previous, parent, true);
                    result = null;           
                      
                }
                else if (repeaterElement.IsIndex)
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
                        var parsedLastElement = this.ProcessElements(repeaterElement.Elements, dataReader, previous, result, index, ref endIfElement, true);
                        if (repeaterElement.Elements.Any(re => re.IsItemTable) || repeaterElement.Elements.Any(re => re.IsItemRepeater && !ItemRepeaterGenerator.CheckInlineWrappingMode(re.StartTag, re.EndTag)))
                        {
                            previous = parsedLastElement;
                            result = null;
                        }

                    }
                    else
                    {
                        element.Value = repeaterElement.XElement.Value;
                    }
                }
                if (result != null)
                {
                    if (!nested)
                    {
                        previous.AddAfterSelf(result);
                        previous = result;
                    }
                    else
                    {
                        parent.Add(result);
                    }
                }
            }            
            return result ?? previous;
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
