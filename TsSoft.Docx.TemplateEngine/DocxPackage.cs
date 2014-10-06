using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine
{
    internal class DocxPackage
    {
        private const string OfficeDocumentRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private Stream docxStream;

        public virtual XDocument DocumentPartXml { get; private set; }

        public DocxPackage()
        {
        }

        public DocxPackage(Stream docxStream)
        {
            this.docxStream = docxStream;
        }

        public virtual DocxPackage Load()
        {
            docxStream.Seek(0, SeekOrigin.Begin);
            using (Package package = Package.Open(docxStream, FileMode.Open, FileAccess.Read))
            {
                new XDocument();
                var docPart = GetDocumentPart(package);
                using (XmlReader reader = XmlReader.Create(docPart.GetStream()))
                {
                    DocumentPartXml = XDocument.Load(reader);
                }
            }
            return this;
        }

        public virtual DocxPackage Save()
        {
            docxStream.Seek(0, SeekOrigin.Begin);
            using (Package package = Package.Open(docxStream, FileMode.Open, FileAccess.ReadWrite))
            {
                var docPart = GetDocumentPart(package);
                var documentStream = docPart.GetStream();
                documentStream.SetLength(0);
                using (var writer = new XmlTextWriter(documentStream, new UTF8Encoding()))//, new CapitalNamesUtf8Encoding()))
                {
                    DocumentPartXml.Save(writer);
                }
                package.Flush();
            }
            return this;
        }

        private PackagePart GetDocumentPart(Package package)
        {
            PackageRelationship relationship = package.GetRelationshipsByType(OfficeDocumentRelType).FirstOrDefault();
            Uri docUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);
            return package.GetPart(docUri);
        }
    }

    internal class CapitalNamesUtf8Encoding : UTF8Encoding
    {
        public override string BodyName { get { return base.BodyName.ToUpper(); } }

        public override string WebName { get { return base.WebName.ToUpper(); } }

        public override string HeaderName { get { return base.HeaderName.ToUpper(); } }
    }
}