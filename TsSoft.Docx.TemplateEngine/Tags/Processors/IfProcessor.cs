using System;
using System.Linq;

namespace TsSoft.Docx.TemplateEngine.Tags.Processors
{
    internal class IfProcessor : AbstractProcessor
    {
        public IfTag Tag { get; set; }

        public override void Process()
        {
            base.Process();
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
                this.CleanUp(this.Tag.StartIf, this.Tag.EndIf);
            }
            else
            {
                this.Tag.StartIf.Remove();
                this.Tag.EndIf.Remove();
            }
        }
    }
}
