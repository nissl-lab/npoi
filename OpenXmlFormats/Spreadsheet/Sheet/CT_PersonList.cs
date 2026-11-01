using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    /// <summary>
    /// Represents the root <c>personList</c> element of the persons part
    /// (namespace http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments).
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments")]
    [XmlRoot(Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
        ElementName = "personList")]
    public class CT_PersonList
    {
        private List<CT_Person> personField = null; // optional field [0..*]

        public static CT_PersonList Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PersonList ctObj = new CT_PersonList();
            ctObj.person = new List<CT_Person>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "person")
                    ctObj.person.Add(CT_Person.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>");
            sw.Write("<personList xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            if (this.person != null)
            {
                foreach (CT_Person x in this.person)
                {
                    x.Write(sw, "person");
                }
            }
            sw.Write("</personList>");
        }

        [XmlElement("person")]
        public List<CT_Person> person
        {
            get { return this.personField; }
            set { this.personField = value; }
        }

        public int SizeOfPersonArray()
        {
            return (null == personField) ? 0 : personField.Count;
        }

        public CT_Person GetPersonArray(int index)
        {
            return personField[index];
        }

        public CT_Person InsertNewPerson(int index)
        {
            if (null == personField) { personField = new List<CT_Person>(); }
            CT_Person ct = new CT_Person();
            personField.Insert(index, ct);
            return ct;
        }

        public CT_Person AddNewPerson()
        {
            if (null == personField) { personField = new List<CT_Person>(); }
            CT_Person ct = new CT_Person();
            personField.Add(ct);
            return ct;
        }

        public void RemovePerson(int index)
        {
            personField.RemoveAt(index);
        }

        public CT_Person[] GetPersonArray()
        {
            if (this.personField == null)
                this.personField = new List<CT_Person>();
            return this.personField.ToArray();
        }
    }
}
