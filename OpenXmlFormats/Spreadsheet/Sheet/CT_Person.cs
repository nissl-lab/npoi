using NPOI.OpenXml4Net.Util;
using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    /// <summary>
    /// Represents a single person in the persons part, the <c>person</c> element
    /// (namespace http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments).
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments")]
    [XmlRoot(Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
        ElementName = "person")]
    public class CT_Person
    {
        private string displayNameField = string.Empty; // required attribute
        private string idField = string.Empty; // required attribute
        private string userIdField = string.Empty; // required attribute
        private string providerIdField = string.Empty; // required attribute

        public static CT_Person Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Person ctObj = new CT_Person();
            ctObj.displayName = XmlHelper.ReadString(node.Attributes["displayName"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["id"]);
            ctObj.userId = XmlHelper.ReadString(node.Attributes["userId"]);
            ctObj.providerId = XmlHelper.ReadString(node.Attributes["providerId"]);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.WriteStart(nodeName);
            XmlHelper.WriteAttribute(sw, "displayName", this.displayName);
            XmlHelper.WriteAttribute(sw, "id", this.id);
            XmlHelper.WriteAttribute(sw, "userId", this.userId);
            XmlHelper.WriteAttribute(sw, "providerId", this.providerId);
            sw.Write("/>");
        }

        /// <summary>
        /// Display name of the person, e.g. "Jane Doe"
        /// </summary>
        [XmlAttribute("displayName")]
        public string displayName
        {
            get { return this.displayNameField; }
            set { this.displayNameField = value; }
        }

        /// <summary>
        /// Id of the person used by threaded comments in this workbook.
        /// </summary>
        [XmlAttribute("id")]
        public string id
        {
            get { return this.idField; }
            set { this.idField = value; }
        }

        /// <summary>
        /// Globally unique user id of the person, stable across workbooks
        /// </summary>
        [XmlAttribute("userId")]
        public string userId
        {
            get { return this.userIdField; }
            set { this.userIdField = value; }
        }

        /// <summary>
        /// Provider the user id belongs to, e.g. "AD"
        /// </summary>
        [XmlAttribute("providerId")]
        public string providerId
        {
            get { return this.providerIdField; }
            set { this.providerIdField = value; }
        }
    }
}
