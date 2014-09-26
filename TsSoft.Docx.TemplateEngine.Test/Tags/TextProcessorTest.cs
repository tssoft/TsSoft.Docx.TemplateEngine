using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    [TestClass]
    public class TextProcessorTest : BaseTagTest
    {
        [TestMethod]
        public void TestMethod1()
        {
            var e = new System.Text.UTF8Encoding();
            Assert.AreEqual("UTF-8", e.BodyName);
            //var asm = Assembly.GetExecutingAssembly().GetManifestResourceStream("TextProcessorTest.docx");
            //ValidateTagsRemoved(null);
        }
    }
}