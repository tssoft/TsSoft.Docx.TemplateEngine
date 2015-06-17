using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TsSoft.Commons.Utils;

namespace TsSoft.Docx.TemplateEngine.Test
{
    [TestClass]
    public class DocxPackageHeaderAndFooterTest
    {
        [TestMethod]
        public void TestLoadDocument()
        {
            using (var stream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageHeaderFooterTest.docx"))
            {
                var package = new DocxPackage(stream).Load();

                var expectedDocument = XDocument.Load(AssemblyResourceHelper.GetResourceStream(this, "DocxPackageHeaderFooterTest_document.xml"));

                var documentPart = package.Parts.Single(p => p.PartUri.OriginalString.Contains("document.xml")).PartXml;

                Assert.AreEqual(expectedDocument.ToString().Trim(), documentPart.Root.ToString().Trim());
            }
        }

        [TestMethod]
        public void TestLoadHeader()
        {
            using (var stream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageHeaderFooterTest.docx"))
            {
                var package = new DocxPackage(stream).Load();

                var expectedHeader = XDocument.Load(AssemblyResourceHelper
                    .GetResourceStream(this, "DocxPackageHeaderFooterTest_header.xml"));

                var headerPart = package.Parts.Single(p => p.PartUri.OriginalString.Contains("header1.xml")).PartXml;

                Assert.AreEqual(expectedHeader.ToString().Trim(), headerPart.Root.ToString().Trim());
            }
        }

        [TestMethod]
        public void TestLoadFooter()
        {
            using (var stream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageHeaderFooterTest.docx"))
            {
                var package = new DocxPackage(stream).Load();

                var expectedFooter = XDocument.Load(AssemblyResourceHelper
                    .GetResourceStream(this, "DocxPackageHeaderFooterTest_footer.xml"));

                var footerPart = package.Parts.Single(p => p.PartUri.OriginalString.Contains("footer1.xml")).PartXml;

                Assert.AreEqual(expectedFooter.ToString().Trim(), footerPart.Root.ToString().Trim());
            }
        }

        [TestMethod]
        public void TestSave()
        {
            using (var docxStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageHeaderFooterTest.docx"))
            {
                var memoryStream = new MemoryStream();
                memoryStream.Seek(0, SeekOrigin.Begin);
                docxStream.Seek(0, SeekOrigin.Begin);
                docxStream.CopyTo(memoryStream);
                var package = new DocxPackage(memoryStream);

                package.Load();
                foreach (var part in package.Parts)
                {
                    XElement sdt = part.PartXml.Descendants(WordMl.SdtName).First();
                    sdt.Remove();
                }
                package.Save();

                var expectedDocumentStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageHeaderFooterTest_saved.docx");

                DocxPackage expectedPackagePart = new DocxPackage(expectedDocumentStream).Load();
                DocxPackage actualPackagePart = new DocxPackage(memoryStream).Load();

                foreach (var expected in expectedPackagePart.Parts)
                {
                    var actual = actualPackagePart.Parts.Single(x => x.PartUri == expected.PartUri);
                    Assert.AreEqual(expected.PartXml.Root.ToString().Trim(), actual.PartXml.Root.ToString().Trim());
                }
            }
        }
    }
}