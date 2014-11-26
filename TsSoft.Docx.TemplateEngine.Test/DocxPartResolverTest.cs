using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TsSoft.Commons.Utils;

namespace TsSoft.Docx.TemplateEngine.Test
{
    [TestClass]
    public class DocxPartResolverTest
    {
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "Reviewed. Suppression is OK here."),TestMethod]
        public void TestCreateAfChunkPart()
        {
            using (var docxStream = AssemblyResourceHelper.GetResourceStream(this, "DocxPackageTest.docx"))
            {
                var memoryStream = new MemoryStream();
                memoryStream.Seek(0, SeekOrigin.Begin);
                docxStream.Seek(0, SeekOrigin.Begin);
                docxStream.CopyTo(memoryStream);
                var testPackage = new DocxPackage(memoryStream);

                const int ExpectedAfChunkId = 10;
                const string ExpectedHtmlString = "<html><head/><body>Test AfChunk Generator</body></html>";
                var expectedPartPath = string.Format("//word/afchunk{0}.dat", ExpectedAfChunkId);
                using (var package = Package.Open(memoryStream, FileMode.Open, FileAccess.ReadWrite))
                {
                    DocxPartResolver.CreateAfChunkPart(package, ExpectedAfChunkId, ExpectedHtmlString);
                    var formattedAltChunkRel = string.Format("altChunkId{0}", ExpectedAfChunkId);
                    var actualDocPart = DocxPartResolver.GetDocumentPart(package);
                    var actualRelationship = actualDocPart.GetRelationship(formattedAltChunkRel);
                    Assert.IsNotNull(actualRelationship);
                    Assert.AreEqual(formattedAltChunkRel, actualRelationship.Id);
                    var actualAfChunkPart = package.GetPart(PackUriHelper.ResolvePartUri(actualRelationship.SourceUri,
                                                                                         actualRelationship.TargetUri));
                    Assert.IsNotNull(actualAfChunkPart);
                    using (var afChunkStream = actualAfChunkPart.GetStream())
                    {
                        using (var reader = new StreamReader(afChunkStream))
                        {
                            Assert.AreEqual(ExpectedHtmlString, reader.ReadToEnd());
                        }
                    }
                }
            }
        }
    }
}
