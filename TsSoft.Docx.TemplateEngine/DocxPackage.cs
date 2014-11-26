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
            using (var package = Package.Open(this.docxStream, FileMode.Open, FileAccess.Read))
            {
                var docPart = DocxPartResolver.GetDocumentPart(package);
                using (var reader = XmlReader.Create(docPart.GetStream()))
                {
                    this.DocumentPartXml = XDocument.Load(reader);
                }
            }
            return this;
        }

        public virtual DocxPackage Save()
        {
            this.docxStream.Seek(0, SeekOrigin.Begin);
            using (var package = Package.Open(this.docxStream, FileMode.Open, FileAccess.ReadWrite))
            {
                var docPart = DocxPartResolver.GetDocumentPart(package);
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

        public virtual void ReplaceAltChunks()
        {
            var document = DocumentPartXml.Document;
            var htmlChunks = HtmlContentProcessor.GenerateAltChunks(document.Root);
            using (var package = Package.Open(this.docxStream, FileMode.Open, FileAccess.ReadWrite))
            {
                for (var i = 1; i <= htmlChunks.Count; i++)
                {
                    DocxPartResolver.CreateAfChunkPart(package, i, htmlChunks[i - 1]);
                }
            }
        }               
    }
}