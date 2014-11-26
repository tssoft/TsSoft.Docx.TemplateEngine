using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace TsSoft.Docx.TemplateEngine
{
    internal class RelationshipMl
    {
        public static readonly XNamespace RelationshipMlNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XName IdName = RelationshipMlNamespace + "id";
    }
}
