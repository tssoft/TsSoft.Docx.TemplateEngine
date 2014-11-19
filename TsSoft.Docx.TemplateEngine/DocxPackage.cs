using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine
{
    internal class DocxPackage
    {
        private const string OfficeDocumentRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string OfficeAfChunkRelType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";

        private Stream docxStream;

        public DocxPackage()
        {
        }

        public DocxPackage(Stream docxStream)
        {
            this.docxStream = docxStream;
        }

        public virtual XDocument DocumentPartXml { get; private set; }

        public virtual DocxPackage Load()
        {
            this.docxStream.Seek(0, SeekOrigin.Begin);
            using (Package package = Package.Open(this.docxStream, FileMode.Open, FileAccess.Read))
            {
                var docPart = this.GetDocumentPart(package);
                using (XmlReader reader = XmlReader.Create(docPart.GetStream()))
                {
                    this.DocumentPartXml = XDocument.Load(reader);
                }
            }
            return this;
        }

        public virtual DocxPackage Save()
        {
            this.docxStream.Seek(0, SeekOrigin.Begin);
            using (Package package = Package.Open(this.docxStream, FileMode.Open, FileAccess.ReadWrite))
            {
                var docPart = this.GetDocumentPart(package);
                var documentStream = docPart.GetStream();
                documentStream.SetLength(0);
                using (var writer = new XmlTextWriter(documentStream, new UTF8Encoding()))
                {
                    this.DocumentPartXml.Save(writer);
                }                
                package.Flush();
            }
            return this;
        }

        public virtual void GenerateAltChunk()
        {
            var document = DocumentPartXml.Document;
            var htmlTags = document.Root.Descendants().Where(el => el.IsTag(HtmlContentProcessor.ProcessedHtmlContentTagName)).ToList();
            var htmlChunks = new List<string>();            
            foreach (var htmlTag in htmlTags)
            {
                var htmlString = htmlTag.GetExpression();
                var htmlChunkIndex = htmlChunks.FindIndex(html => html.Equals(htmlString));
                if (htmlChunkIndex == -1)
                {
                    htmlChunks.Add(htmlString);
                    htmlChunkIndex = htmlChunks.Count - 1;
                    using (var package = Package.Open(this.docxStream, FileMode.Open, FileAccess.ReadWrite))
                    {
                        this.CreateAfChunkPart(package, htmlChunkIndex + 1, htmlString);
                    }
                }
                var parent = htmlTag.Parent;
                var altChunkElement = DocxHelper.CreateAltChunkElement(htmlChunkIndex + 1);
                if (!parent.Name.Equals(WordMl.ParagraphName))
                {
                    htmlTag.AddAfterSelf(altChunkElement);
                }
                else
                {
                    parent.AddAfterSelf(altChunkElement);
                    if (parent.Elements().Count(el => !el.Name.Equals(WordMl.ParagraphPropertiesName)) == 1)
                    {
                        parent.Remove();
                    }
                }
                htmlTag.Remove();
            }            
        }

        private PackagePart GetDocumentPart(Package package)
        {
            PackageRelationship relationship = package.GetRelationshipsByType(OfficeDocumentRelType).FirstOrDefault();                        
            Uri docUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);                        
            return package.GetPart(docUri);
        }

        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1305:FieldNamesMustNotUseHungarianNotation", Justification = "Reviewed. Suppression is OK here.")]
        private void CreateAfChunkPart(Package package, int afChunkId, string htmlString)
        {
            var formattedAfChunk = string.Format("afchunk{0}", afChunkId);
            var formattedAfChunkPartPath = string.Format("word/{0}.dat", formattedAfChunk);
            var docPart = this.GetDocumentPart(package);            
            var partUriAfChunk = PackUriHelper.CreatePartUri(new Uri(formattedAfChunkPartPath, UriKind.Relative));
            var packagePartAfChunk = package.CreatePart(partUriAfChunk, "application/xhtml+xml");
            using (var afChunkStream = packagePartAfChunk.GetStream())
            {
                using (var stringStream = new StreamWriter(afChunkStream))
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