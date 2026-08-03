using System.IO;
using System.Xml;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    /// <summary>
    /// Document wrapper around the persons part of a workbook.
    /// </summary>
    public class PersonListDocument
    {
        private CT_PersonList personList = null;

        public PersonListDocument()
        {
        }

        public PersonListDocument(CT_PersonList personList)
        {
            this.personList = personList;
        }

        public static PersonListDocument Parse(XmlDocument xmlDoc, XmlNamespaceManager namespaceManager)
        {
            PersonListDocument doc = new PersonListDocument();
            doc.personList = CT_PersonList.Parse(xmlDoc.DocumentElement, namespaceManager);
            return doc;
        }

        public CT_PersonList GetPersonList()
        {
            return personList;
        }

        public void SetPersonList(CT_PersonList personList)
        {
            this.personList = personList;
        }

        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                this.personList.Write(sw);
            }
        }
    }
}
