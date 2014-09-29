using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine.Test
{
    internal class A
    {
        public int MyProperty { get; set; }
    }

    [TestClass]
    public class DocxGeneratorTest
    {
        [TestMethod]
        public void TestFillDocx()
        {
            //c:\Temp\document.xml
            var doc = XDocument.Load(@"c:\Temp\document.xml");
            XNamespace ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var n = doc.Descendants(ns + "sdt").Skip(1).First();
            //.Descendants(ns + "tag").Attributes(ns + "val").First().Value
            Assert.AreEqual(100,  n.DescendantNodesAndSelf());


            //var generator = new DocxGenerator<object>();
        }
    }
}