using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TsSoft.Docx.TemplateEngine
{
    public class DocxGeneratorSettings
    {
        public MissingDataMode MissingDataMode { get; set; }

        public bool LockDynamicContent { get; set; }
    }

    internal class MissingDataException : Exception
    {
        public MissingDataException(string missingPath)
            : base(string.Format(MessageStrings.DataNotFound, missingPath))
        {
            
        }
    }
}
