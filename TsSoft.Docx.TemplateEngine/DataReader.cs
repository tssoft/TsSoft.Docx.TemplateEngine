using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Xml.XPath;

namespace TsSoft.Docx.TemplateEngine
{
    internal class DataReader
    {
        private readonly XElement rootElement;
        private const string pathArgumentName = "path";

        public DataReader()
        {
        }

        public DataReader(XElement rootElement)
        {
            this.rootElement = rootElement;
        }

        public string ReadText(string expression)
        {
            var textElement = rootElement.XPathSelectElement(expression);
            return textElement.Value;
        }

        public DataReader GetReader(string path)
        {
            if (path == null)
            {
                throw new ArgumentNullException(pathArgumentName);
            }

            var newElement = rootElement.XPathSelectElement(path);

            return newElement != null ? new DataReader(newElement) : null;
        }

        public virtual IEnumerable<DataReader> GetReaders(string path)
        {
            if (path == null)
            {
                throw new ArgumentNullException(pathArgumentName);
            }

            var newElements = rootElement.XPathSelectElements(path);

            foreach (var element in newElements)
            {
                yield return new DataReader(element);
            }
        }

        public override bool Equals(Object obj)
        {
            if (!(obj is DataReader))
            {
                return false;
            }

            DataReader other = (DataReader)obj;

            bool result;
            if (this == null && other == null)
            {
                result = true;
            }
            else
            {
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

                result = XElement.DeepEquals(thisRootElement, otherRootElement);
            }
            return result;
        }
    }
}