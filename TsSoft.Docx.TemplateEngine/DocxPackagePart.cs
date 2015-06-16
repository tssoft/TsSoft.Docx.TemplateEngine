using System;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using TsSoft.Docx.TemplateEngine.Tags.Processors;

namespace TsSoft.Docx.TemplateEngine
{
    public class DocxPackagePart
    {  
        private readonly Stream docxStream;

        public DocxPackagePart()
        {
        }

        public DocxPackagePart(Stream docxStream)
        {
            this.docxStream = docxStream;
        }
        public DocxPackagePart(Stream docxStream, Uri partUri, XDocument partXml)
        {
            this.docxStream = docxStream;

            PartUri = partUri;
            PartXml = partXml;
        }

        public virtual Uri PartUri { get; private set; }

        public virtual XDocument PartXml { get; private set; }

        public virtual DocxPackagePart Save()
        {
            this.docxStream.Seek(0, SeekOrigin.Begin);
            using (var package = Package.Open(this.docxStream, FileMode.Open, FileAccess.ReadWrite))
            {
                var part = DocxPartResolver.GetPart(package, PartUri);
                var stream = part.GetStream();
                stream.SetLength(0);
                using (var writer = new XmlTextWriter(stream, new UTF8Encoding()))
                {
                    this.PartXml.Save(writer);
                }                
                package.Flush();
            }
            return this;
        }

        public virtual void ReplaceAltChunks()
        {
            var document = PartXml.Document;
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