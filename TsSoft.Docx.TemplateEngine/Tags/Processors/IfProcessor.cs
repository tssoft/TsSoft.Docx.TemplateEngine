using System;
using System.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    using System.Xml.Linq;

    internal class IfProcessor : AbstractProcessor
    {
        public IfTag Tag { get; set; }

        public override void Process()
        {
            base.Process();
            this.ProcessDynamicContent();

            bool truthful;
            try
            {
                truthful = bool.Parse(this.DataReader.ReadText(this.Tag.Conidition));
            }
            catch (System.FormatException)
            {
                truthful = false;
            }
            catch (System.Xml.XPath.XPathException)
            {
                truthful = false;
            }
            if (!truthful)
            {
                if (this.LockDynamicContent)
                {
                    this.Tag.StartIf.AddBeforeSelf(DocxHelper.CreateDynamicContentElement(Enumerable.Empty<XElement>(), this.Tag.StartIf));
                }
                this.CleanUp(this.Tag.StartIf, this.Tag.EndIf);
            }
            else
            {
                if (this.LockDynamicContent)
                {
                    var innerElements = TraverseUtils.ElementsBetween(this.Tag.StartIf, this.Tag.EndIf).ToList();
                    innerElements.Remove();
                    this.Tag.StartIf.AddBeforeSelf(DocxHelper.CreateDynamicContentElement(innerElements, this.Tag.StartIf));
                    this.CleanUp(this.Tag.StartIf, this.Tag.EndIf);
                }
                else
                {
                    //this.CleanUp(Tag.StartIf, Tag.EndIf);
                    this.Tag.StartIf.Remove();
                    this.Tag.EndIf.Remove();
                }
            }
        }

        private void ProcessDynamicContent()
        {
            var dynamicContentTags =
                TraverseUtils.ElementsBetween(this.Tag.StartIf, this.Tag.EndIf).Where(element => element.IsSdt()).ToList();
            foreach (var dynamicContentTag in dynamicContentTags)
            {
                var innerElements = dynamicContentTag.Element(WordMl.SdtContentName).Elements();
                dynamicContentTag.AddAfterSelf(innerElements);
                dynamicContentTag.Remove();
            }
        }
    }
}
