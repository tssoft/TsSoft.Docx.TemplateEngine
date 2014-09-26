using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TsSoft.Docx.TemplateEngine.Test
{
    [TestClass]
    public class DocxGeneratorTest
    {
        [TestMethod]
        public void TestFillDocx()
        {
            var generator = new DocxGenerator<object>();
        }
    }
}
