using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;

namespace TsSoft.Docx.TemplateEngine
{
    internal class DocxPartResolver
    {
        private const string OfficeDocumentRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string HeaderRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
        private const string FooterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
        private const string OfficeAfChunkRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";

        public static PackagePart GetDocumentPart(Package package)
        {
            PackageRelationship relationship = package.GetRelationshipsByType(OfficeDocumentRelType).FirstOrDefault();
            Uri docUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);
            return package.GetPart(docUri);
        }

        public static IEnumerable<PackagePart> GetHeaderParts(Package package)
        {
            var documentPart = GetDocumentPart(package);
            var headerRelationships = documentPart.GetRelationshipsByType(HeaderRelationshipType);
            foreach (var relationship in headerRelationships)
            {
                yield return package.GetPart(PackUriHelper.ResolvePartUri(documentPart.Uri, relationship.TargetUri));
            }
        }

        public static IEnumerable<PackagePart> GetFooterParts(Package package)
        {
            var documentPart = GetDocumentPart(package);
            var footerRelationships = documentPart.GetRelationshipsByType(FooterRelationshipType);
            foreach (var relationship in footerRelationships)
            {
                yield return package.GetPart(PackUriHelper.ResolvePartUri(documentPart.Uri, relationship.TargetUri));
            }
        }

        public static PackagePart GetPart(Package package, Uri uri)
        {
            return package.GetPart(uri);
        }

        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "Reviewed. Suppression is OK here.")]
        public static void CreateAfChunkPart(Package package, int afChunkId, string htmlString)
        {
            const string HtmlContentType = "application/xhtml+xml";
            var formattedAfChunk = string.Format("afchunk{0}", afChunkId);
            var formattedAfChunkPartPath = string.Format("word/{0}.dat", formattedAfChunk);
            var docPart = GetDocumentPart(package);
            var partUriAfChunk = PackUriHelper.CreatePartUri(new Uri(formattedAfChunkPartPath, UriKind.Relative));
            var packagePartAfChunk = package.CreatePart(partUriAfChunk, HtmlContentType);            
            using (var afChunkStream = packagePartAfChunk.GetStream())
            {
                using (var stringStream = new StreamWriter(afChunkStream, Encoding.UTF8))
                {
                    stringStream.Write(htmlString);
                }
            }
            docPart.CreateRelationship(
                                       new Uri(formattedAfChunk + ".dat", UriKind.Relative),
                                       TargetMode.Internal,
                                       OfficeAfChunkRelType,
                                       string.Format("altChunkId{0}", afChunkId));
        }
    }
}
