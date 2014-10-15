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
        private MissingDataMode missingDataModeSettings;

        public DataReader()
        {
           
        }
        
        public DataReader(XElement rootElement)
        {
            this.rootElement = rootElement;
            /*
            foreach (var node in rootElement.DescendantsAndSelf())
            {
                node.Name = node.Name.ToString().ToLower();
            }
             * */
            //Console.WriteLine(rootElement.ToString());
        }
        
        public MissingDataMode MissingDataModeSettings
        {
            get { return this.missingDataModeSettings; }
            set { this.missingDataModeSettings = value; }            
        }

        public string ReadText(string expression)
        {
            var textElement = this.rootElement.XPathSelectElement(expression.ToLower());
            if (textElement == null)
            {
                switch (this.missingDataModeSettings)
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
                dataReader.MissingDataModeSettings = this.missingDataModeSettings;
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
                if (this.missingDataModeSettings == MissingDataMode.ThrowException)
                {
                    throw new MissingDataException(path);
                }
                return Enumerable.Empty<DataReader>();
            }
            var newElements = this.rootElement.XPathSelectElements(path.ToLower());
            IEnumerable<DataReader> readers = newElements.Select(element => new DataReader(element)).ToList();            
            foreach (var dataReader in readers)
            {
                dataReader.MissingDataModeSettings = this.missingDataModeSettings;
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