using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Test.Tags.Processors
{
    public class BaseProcessorTest
    {
        /// <summary>
        /// Check for number of tags is 0
        /// </summary>
        /// <param name="doc"></param>
        /// <author>Георгий Поликарпов</author>
        protected void ValidateTagsRemoved(XmlDocument doc)
        {
            XDocument document = XDocument.Load(new XmlNodeReader(doc));
            IEnumerable<XElement> tags = document.Descendants(WordMl.SdtName);
            Assert.IsFalse(tags.Any());
        }
    }
}