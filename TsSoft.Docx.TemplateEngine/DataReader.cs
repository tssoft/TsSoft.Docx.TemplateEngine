using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Xml.XPath;

namespace TsSoft.Docx.TemplateEngine
{
    using System.Linq;

    internal class DataReader
    {
        private const string PathArgumentName = "path";
        private readonly XElement rootElement;

        public DataReader()
        {
        }

        public DataReader(XElement rootElement)
        {
            this.rootElement = rootElement;
        }

        public string ReadText(string expression)
        {
            var textElement = this.rootElement.XPathSelectElement(expression);
            return textElement.Value;
        }

        public DataReader GetReader(string path)
        {
            if (path == null)
            {
                throw new ArgumentNullException(PathArgumentName);
            }

            var newElement = this.rootElement.XPathSelectElement(path);

            return newElement != null ? new DataReader(newElement) : null;
        }

        public virtual IEnumerable<DataReader> GetReaders(string path)
        {
            if (path == null)
            {
                throw new ArgumentNullException(PathArgumentName);
            }

            var newElements = this.rootElement.XPathSelectElements(path);

            return newElements.Select(element => new DataReader(element));
        }

        public override bool Equals(object obj)
        {
            if (!(obj is DataReader))
            {
                return false;
            }

            var other = (DataReader)obj;

            var thisRootElement = this.rootElement;
            if (thisRootElement != null)
            {
                thisRootElement.RemoveAttributes();
            }

            var otherRootElement = other.rootElement;
            if (otherRootElement != null)
            {
                otherRootElement.RemoveAttributes();
            }

            return XNode.DeepEquals(thisRootElement, otherRootElement);
        }
    }
}