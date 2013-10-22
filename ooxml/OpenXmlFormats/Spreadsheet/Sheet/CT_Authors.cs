using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    // [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    // not needed because it not used as a root [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", ElementName = "authors")]
    public class CT_Authors
    {

        private List<string> authorField = null; // optional field [0..*]

        //public CT_Authors()
        //{
        //    this.authorField = new List<string>();
        //}
        public int SizeOfAuthorArray()
        {
            return (null == authorField) ? 0 : authorField.Count;
        }
        public string GetAuthorArray(int index)
        {
            return (null == authorField) ? null : authorField[index];
        }
        public void Insert(int index, string author)
        {
            if (null == authorField) { authorField = new List<string>(); }
            authorField.Insert(index, author);
        }
        public void AddAuthor(string name)
        {
            if (null == authorField) { authorField = new List<string>(); }
            authorField.Add(name);
        }
        //[XmlArray("authors", Order = 0)] // - encapsulates the following items, but the outer element already provides the container.
        //[XmlArrayItem("author")]
        [XmlElement("author")] // this is serialized into multiple author entries
        public List<string> author
        {
            get
            {
                return this.authorField;
            }
            set
            {
                this.authorField = value;
            }
        }
    }
}
