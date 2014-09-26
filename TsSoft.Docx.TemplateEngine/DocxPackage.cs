using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;

namespace TsSoft.Docx.TemplateEngine
{
    internal class DocxPackage
    {
        private const string OfficeDocumentRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public const string WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private Stream docxStream;

        public XmlDocument documentPartXml { get; private set; }

        public DocxPackage(Stream docxStream)
        {
            this.docxStream = docxStream;
        }

        public void Load()
        {
            docxStream.Seek(0, SeekOrigin.Begin);
            using (Package package = Package.Open(docxStream, FileMode.Open, FileAccess.Read))
            {
                var nameTable = new NameTable();
                var nsManager = new XmlNamespaceManager(nameTable);
                nsManager.AddNamespace("w", WordMlNamespace);
                documentPartXml = new XmlDocument(nameTable);
                var docPart = GetDocumentPart(package);
                documentPartXml.Load(docPart.GetStream());
            }
        }

        public void Save()
        {
            docxStream.Seek(0, SeekOrigin.Begin);
            using (Package package = Package.Open(docxStream, FileMode.Open, FileAccess.Read))
            {
                var docPart = GetDocumentPart(package);
                var documentStream = docPart.GetStream();
                documentStream.SetLength(documentPartXml.InnerXml.Length);
                using (var writer = new XmlTextWriter(documentStream, new CapitalNamesUtf8Encoding()))
                {
                    //writer.Formatting = Formatting.Indented;
                    documentPartXml.Save(writer);
                }
                package.Flush();
            }
        }

        private PackagePart GetDocumentPart(Package package)
        {
            PackageRelationship relationship = package.GetRelationshipsByType(OfficeDocumentRelType).FirstOrDefault();
            Uri docUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);
            return package.GetPart(docUri);
        }
    }

    internal class CapitalNamesUtf8Encoding : System.Text.UTF8Encoding
    {
        public override string BodyName { get { return base.BodyName.ToUpper(); } }

        public override string WebName { get { return base.WebName.ToUpper(); } }

        public override string HeaderName { get { return base.HeaderName.ToUpper(); } }
    }
}