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
                XElement endIfElementTmp = null;
                current = this.ProcessElements(repeaterElements, dataReaders[index], current, null, index + 1, ref endIfElementTmp);
            }
            foreach (var repeaterElement in repeaterElements)
            {
                // TODO one
                /*if (repeaterElement.HasElements &&
                    repeaterElement.Elements.Any(re => ItemRepeaterGenerator.IsItemRepeaterElement(re.XElement)))
                {
                    continue;
                }*/
                repeaterElement.XElement.Remove();
                
            }

            if (this.LockDynamicContent)
            {
                var innerElements =
                    TraverseUtils.ElementsBetween(this.RepeaterTag.StartRepeater, this.RepeaterTag.EndRepeater).ToList();
                innerElements.Remove();
                this.RepeaterTag.StartRepeater.AddBeforeSelf(DocxHelper.CreateDynamicContentElement(innerElements,
                                                                                                    this.RepeaterTag
                                                                                                        .StartRepeater));
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

        private XElement ProcessElements(IEnumerable<RepeaterElement> elements, DataReader dataReader, XElement start, XElement parent, int index, ref XElement endIfElement)
        {
            XElement result = null;
            XElement previous = start; 
            elements = elements.Where(el => !el.XElement.IsTag(ItemRepeaterTags.ItemTag) && !el.XElement.IsTag(ItemRepeaterTags.ItemIf) && !el.XElement.IsTag(ItemRepeaterTags.EndItemIf) && !el.XElement.IsTag(ItemRepeaterTags.EndItemRepeaterTagName) && !el.XElement.IsTag(ItemRepeaterTags.IndexTag));            
            foreach (var repeaterElement in elements)
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
                if (repeaterElement.IsItemHtmlContent)
                {
                    result = HtmlContentProcessor.MakeHtmlContentProcessed(repeaterElement.XElement, dataReader.ReadText(repeaterElement.Expression), true);                    
                }
                else if (repeaterElement.IsItemRepeater)
                {                                        
                    var itemRepeaterTag = new ItemRepeaterTag()
                        {
                            StartItemRepeater = repeaterElement.StartTag,
                            EndItemRepeater = repeaterElement.EndTag,
                            Source = repeaterElement.Expression                            
                        };                    
                   // var itemRepeaterParser = new ItemRepeaterParser();                    
                    //previous = itemRepeaterParser.Parse(itemRepeaterTag, dataReader.GetReaders(itemRepeaterTag.Source).ToList(), previous);
                    var itemRepeaterGenerator = new ItemRepeaterGenerator();
                    previous = itemRepeaterGenerator.Generate(itemRepeaterTag,
                                                              dataReader.GetReaders(repeaterElement.Expression),
                                                              previous, parent, true);
                    result = null;
                    continue;                    
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
                        
                        var parsedLastElement = this.ProcessElements(repeaterElement.Elements, dataReader, null, result, index, ref endIfElement);
                        if (repeaterElement.Elements.Any(re => re.IsItemRepeater))
                        {
                            previous = parsedLastElement;
                        }

                    }
                    else
                    {
                        element.Value = repeaterElement.XElement.Value;
                    }
                }
                if (result != null)
                {
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
