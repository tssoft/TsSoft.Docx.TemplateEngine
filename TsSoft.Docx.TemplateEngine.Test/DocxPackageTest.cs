using System;
using System.IO.Packaging;

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
            using(var docxStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageTest.docx"))
            {
                var testPackage = new DocxPackage(docxStream);

                testPackage.Load();

                Stream expectedDocStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageTest_document.xml");
                var expectedDocument = XDocument.Load(expectedDocStream);

                var documentPart = testPackage.Parts.Single(p => p.PartUri.OriginalString.Contains("document.xml")).PartXml;

                Assert.AreEqual(expectedDocument.ToString().Trim(), documentPart.Root.ToString().Trim());
            }
        }

        [TestMethod]
        public void TestSave()
        {
            using (var docxStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageTest.docx"))
            {
                var memoryStream = new MemoryStream();
                memoryStream.Seek(0, SeekOrigin.Begin);
                docxStream.Seek(0, SeekOrigin.Begin);
                docxStream.CopyTo(memoryStream);
                var testPackage = new DocxPackage(memoryStream);

                testPackage.Load();

                var documentPart = testPackage.Parts.Single(p => p.PartUri.OriginalString.Contains("document.xml")).PartXml;
                XElement sdt = documentPart.Descendants(WordMl.SdtName).First();
                sdt.Remove();

                testPackage.Save();

                var expectedDocumentStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageTest_save.docx");

                DocxPackage expectedPackagePart = new DocxPackage(expectedDocumentStream).Load();
                DocxPackage actualPackagePart = new DocxPackage(memoryStream).Load();
                Assert.AreEqual(expectedPackagePart.Parts.Single(p => p.PartUri.OriginalString.Contains("document.xml")).PartXml.Root.ToString().Trim(),
                                actualPackagePart.Parts.Single(p => p.PartUri.OriginalString.Contains("document.xml")).PartXml.Root.ToString().Trim());
            }
        }                
    }
}