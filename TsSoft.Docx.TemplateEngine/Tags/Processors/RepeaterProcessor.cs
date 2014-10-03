
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    internal class RepeaterProcessor : AbstractProcessor
    {
        public void Do(RepeaterTag repeaterTag)
        {
            var current = repeaterTag.Start;
            var dataReaders = DataReader.GetReaders(repeaterTag.Source).DefaultIfEmpty().ToList();
            for (var index = 0; index < dataReaders.Count; index++)
            {
                current = ProcessElements(repeaterTag.Content, dataReaders[index], current, null, index + 1);
            }
            foreach (var repeaterElement in repeaterTag.Content)
            {
                repeaterElement.XElement.Remove();
            }
            repeaterTag.Start.Remove();
            repeaterTag.End.Remove();
        }

        private XElement ProcessElements(IEnumerable<RepeaterElement> elements, DataReader dataReader, XElement start, XElement parent, int index)
        {
            XElement result = null;
            XElement previous = start;
            foreach (var repeaterElement in elements)
            {
                if (repeaterElement.IsIndex)
                {
                    result = DocxHelper.CreateTextElement(index.ToString(CultureInfo.CurrentCulture));
                }
                else if (repeaterElement.IsItem && repeaterElement.HasExpression)
                {
                    result = DocxHelper.CreateTextElement(dataReader.ReadText(repeaterElement.Expression));
                }
                else
                {
                    var xElement = new XElement(repeaterElement.XElement);
                    xElement.RemoveNodes();
                    result = xElement;
                    if (repeaterElement.HasElements)
                    {
                        ProcessElements(repeaterElement.Elements, dataReader, null, result, index);
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
