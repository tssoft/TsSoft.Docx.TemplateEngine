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

        public bool IsLast { get; private set; }        

        public MissingDataMode MissingDataModeSettings { get; set; }

        public string ReadText(string expression)
        {
            var textElement = this.rootElement.XPathSelectElement(expression.ToLower());
            if (textElement == null)
            {
                switch (this.MissingDataModeSettings)
                {
                    case MissingDataMode.Ignore:
                        return " ";
                    case MissingDataMode.PrintError:
                        return string.Format("'{0}' not found or empty", expression);
                    case MissingDataMode.ThrowException:
                        throw new MissingDataException(expression);                                           
                }
            }                        
            return textElement.Value;            
        }

        public string ReadAttribute(string expression, string attributeName)
        {
            var textElement = this.rootElement.XPathSelectElement(expression.ToLower());
            if (textElement != null)
            {
                var attribute = textElement.Attribute(attributeName);
                return attribute != null ? attribute.Value : string.Empty;
            }
            return string.Empty;
        }

        public DataReader GetReader(string path)
        {
            if (path == null)
            {
                throw new ArgumentNullException(PathArgumentName);
            }

            var newElement = this.rootElement.XPathSelectElement(path.ToLower());

            if (newElement != null)
            {
                var dataReader = new DataReader(newElement);
                dataReader.MissingDataModeSettings = this.MissingDataModeSettings;
                dataReader.IsLast = true;
                return dataReader;
            }
            return null;
        }

        public virtual IEnumerable<DataReader> GetReaders(string path)
        {
            if (path == null)
            {
                throw new ArgumentNullException(PathArgumentName);
            }
            if (this.rootElement.XPathSelectElement(path.ToLower()) == null)
            {
                if (this.MissingDataModeSettings == MissingDataMode.ThrowException)
                {
                    throw new MissingDataException(path);
                }
                return Enumerable.Empty<DataReader>();
            }
            var newElements = this.rootElement.XPathSelectElements(path.ToLower());
            IEnumerable<DataReader> readers = newElements.Select(element => new DataReader(element)).ToList();
            readers.Last().IsLast = true;
            foreach (var dataReader in readers)
            {
                dataReader.MissingDataModeSettings = this.MissingDataModeSettings;
            }
            return readers;
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