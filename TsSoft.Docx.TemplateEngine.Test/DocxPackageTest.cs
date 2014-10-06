namespace TsSoft.Docx.TemplateEngine.Test
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.IO;
    using System.Linq;
    using System.Xml.Linq;
    using TsSoft.Commons.Utils;

    [TestClass]
    public class DocxPackageTest
    {
        [TestMethod]
        public void TestLoad()
        {
            var docxStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageTest.docx");
            var testPackage = new DocxPackage(docxStream);

            testPackage.Load();

            Stream expectedDocStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageTest_document.xml");
            var expectedDocument = XDocument.Load(expectedDocStream);

            Assert.AreEqual(expectedDocument.ToString().Trim(), testPackage.DocumentPartXml.Root.ToString().Trim());
        }

        [TestMethod]
        public void TestSave()
        {
            var docxStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageTest.docx");
            var memoryStream = new MemoryStream();
            memoryStream.Seek(0, SeekOrigin.Begin);
            docxStream.Seek(0, SeekOrigin.Begin);
            docxStream.CopyTo(memoryStream);
            var testPackage = new DocxPackage(memoryStream);

            testPackage.Load();

            XElement sdt = testPackage.DocumentPartXml.Descendants(WordMl.SdtName).First();
            sdt.Remove();

            testPackage.Save();

            var expectedDocumentStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageTest_save.docx");

            DocxPackage expectedPackage = new DocxPackage(expectedDocumentStream).Load();
            DocxPackage actualPackage = new DocxPackage(memoryStream).Load();
            Assert.AreEqual(expectedPackage.DocumentPartXml.Root.ToString().Trim(), actualPackage.DocumentPartXml.Root.ToString().Trim());

        }
    }
}