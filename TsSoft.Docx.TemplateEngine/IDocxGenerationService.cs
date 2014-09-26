using System.IO;
using System.Runtime.Serialization;

namespace TsSoft.Docx.TemplateEngine
{
    public interface IDocxGenerationService<E>
    {
        void GenerateDocx(Stream templateStream, Stream outputStream, E entity);
    }
}