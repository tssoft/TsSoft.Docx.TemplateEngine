using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TsSoft.Docx.TemplateEngine.Test.Tags
{
    using System.Xml.Linq;

    using TsSoft.Commons.Utils;

    [TestClass]
    public class TraverseUtilsTest
    {
        [TestMethod]
        public void TestNextElements()
        {
            var docStream = AssemblyResourceHelper.GetResourceStream(this, "ComplexIf.xml");
            var doc = XDocument.Load(docStream);
            var documentRoot = doc.Root.Element(WordMl.WordMlNamespace + "body");

        }
    }
}
