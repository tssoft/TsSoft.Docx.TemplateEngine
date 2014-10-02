using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Xml.Linq;
using TsSoft.Commons.Utils;

namespace TsSoft.Docx.TemplateEngine.Test
{
    [TestClass]
    public class DocxPackageTest
    {
        [TestMethod]
        public void TestLoadSave()
        {
            using (var stream = new FileStream(@"c:\temp\xxx.docx", FileMode.OpenOrCreate))//new MemoryStream())
            using (var docStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageTest.docx"))
            {
                docStream.CopyTo(stream);
                var pkg = new DocxPackage(stream);
                pkg.Load();
                Assert.IsNotNull(pkg.DocumentPartXml);
                pkg.DocumentPartXml.AddFirst(new XComment("Hellol, world!"));
                pkg.Save();
            }
        }
    }
}