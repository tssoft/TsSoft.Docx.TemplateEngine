using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine
{
    public class DocxPackage
    {
        private Stream docxStream;

        public virtual IList<DocxPackagePart> Parts { get; private set; }

        public DocxPackage()
        {
        }

        public DocxPackage(Stream docxStream)
        {
            this.docxStream = docxStream;
            Parts = new List<DocxPackagePart>();
        }

        public virtual DocxPackage Load()
        {
            this.docxStream.Seek(0, SeekOrigin.Begin);
            using (var package = Package.Open(this.docxStream, FileMode.Open, FileAccess.Read))
            {
                LoadPackageParts(DocxPartResolver.GetDocumentPart(package));
                LoadPackageParts(DocxPartResolver.GetHeaderParts(package).ToArray());
                LoadPackageParts(DocxPartResolver.GetFooterParts(package).ToArray());
            }
            return this;
        }

        private void LoadPackageParts(params PackagePart[] packageParts)
        {
            foreach (var packagePart in packageParts)
            {
                XDocument partXml;
                using (var reader = XmlReader.Create(packagePart.GetStream()))
                    partXml = XDocument.Load(reader);

                this.Parts.Add(new DocxPackagePart(docxStream, new Uri(packagePart.Uri.OriginalString, UriKind.Relative), partXml));
            }
        }

        public virtual void Save()
        {
            foreach (var packagePart in Parts)
            {
                packagePart.Save();
            }
        }
    }
}