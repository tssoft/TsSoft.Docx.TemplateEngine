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
                if (this.DynamicContentMode == DynamicContentMode.Lock)
                {
                    this.Tag.StartIf.AddBeforeSelf(DocxHelper.CreateDynamicContentElement(Enumerable.Empty<XElement>(), this.Tag.StartIf));
                }
                this.CleanUp(this.Tag.StartIf, this.Tag.EndIf);
            }
            else
            {
                switch (this.DynamicContentMode)
                {
                    case DynamicContentMode.NoLock:
                        this.Tag.StartIf.Remove();
                        this.Tag.EndIf.Remove();
                        break;
                    case DynamicContentMode.Lock:
                        var innerElements = TraverseUtils.ElementsBetween(this.Tag.StartIf, this.Tag.EndIf);
                        this.Tag.StartIf.AddBeforeSelf(DocxHelper.CreateDynamicContentElement(innerElements, this.Tag.StartIf));
                        this.CleanUp(this.Tag.StartIf, this.Tag.EndIf);
                        break;
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
