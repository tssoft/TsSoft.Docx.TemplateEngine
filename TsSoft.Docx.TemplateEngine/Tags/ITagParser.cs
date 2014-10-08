using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine.Tags
{
    internal interface ITagParser
    {
        XElement Parse(ITagProcessor parentProcessor, XElement startElement);
    }
}