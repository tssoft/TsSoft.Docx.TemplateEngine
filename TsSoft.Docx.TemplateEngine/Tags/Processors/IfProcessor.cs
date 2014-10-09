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
            bool.TryParse(this.DataReader.ReadText(this.Tag.Conidition), out truthful);
            var body = this.Tag.EndIf.Ancestors().First();
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
