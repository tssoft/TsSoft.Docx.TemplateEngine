namespace TsSoft.Docx.TemplateEngine.Test.Tags.Processors
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Linq;
    using System.Xml.Linq;

    public class BaseProcessorTest
    {
        protected void ValidateTagsRemoved(XContainer document)
        {
            Assert.IsFalse(document.Descendants(WordMl.SdtName).Any());
        }

    }
}