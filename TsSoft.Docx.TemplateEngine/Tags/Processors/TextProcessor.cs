using System.Linq;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{

    internal class TextProcessor : AbstractProcessor
    {
        public TextTag TextTag { get; set; }

        public override void Process()
        {
            base.Process();
            var element = TextTag.TagNode;
            var text = DataReader.ReadText(TextTag.Expression);
            XElement parent = element.Parent;
            if (parent.Name == WordMl.TableRowName)
            {
                parent = element.Element(WordMl.SdtContentName).Element(WordMl.TableCellName);
            }
            var textElement = DocxHelper.CreateTextElement(element, parent, text);
            var result = this.LockDynamicContent
                             ? DocxHelper.CreateDynamicContentElement(new[] { textElement }, this.TextTag.TagNode)
                             : textElement;
            if (element.Parent.Name != WordMl.TableRowName)
            {
                element.AddBeforeSelf(result);
            }
            else
            {
                var cell = new XElement(WordMl.TableCellName, result);    
                cell.AddFirst(element.Descendants(WordMl.TableCellPropertiesName));
                element.AddBeforeSelf(cell);
            }            
            element.Remove();
        }
    }
}