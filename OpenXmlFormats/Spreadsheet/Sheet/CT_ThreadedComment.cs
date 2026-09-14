using NPOI.OpenXml4Net.Util;
using System;
using System.Globalization;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    /// <summary>
    /// Represents a single threaded comment, the <c>threadedComment</c> element of the
    /// <c>ThreadedComments</c> part
    /// (namespace http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments).
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments")]
    [XmlRoot(Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
        ElementName = "threadedComment")]
    public class CT_ThreadedComment
    {
        private string refField = string.Empty; // required attribute
        private DateTime? dTField = null; // required attribute
        private string personIdField = string.Empty; // required attribute
        private string idField = string.Empty; // required attribute
        private string parentIdField = null; // optional attribute
        private string textField = null; // optional element

        public static CT_ThreadedComment Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ThreadedComment ctObj = new CT_ThreadedComment();
            ctObj.@ref = XmlHelper.ReadString(node.Attributes["ref"]);
            ctObj.dT = XmlHelper.ReadDateTime(node.Attributes["dT"]);
            ctObj.personId = XmlHelper.ReadString(node.Attributes["personId"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["id"]);
            ctObj.parentId = XmlHelper.ReadString(node.Attributes["parentId"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "text")
                    ctObj.text = childNode.InnerText;
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.WriteStart(nodeName);
            XmlHelper.WriteAttribute(sw, "ref", this.@ref);
            if (this.dT.HasValue)
                XmlHelper.WriteAttribute(sw, "dT", FormatDateTime(this.dT.Value));
            XmlHelper.WriteAttribute(sw, "personId", this.personId);
            XmlHelper.WriteAttribute(sw, "id", this.id);
            XmlHelper.WriteAttribute(sw, "parentId", this.parentId);
            sw.Write('>');
            if (this.text != null)
            {
                sw.Write("<text>");
                sw.Write(XmlHelper.EncodeXml(this.text));
                sw.Write("</text>");
            }
            sw.WriteEndElement(nodeName);
        }

        private static string FormatDateTime(DateTime value)
        {
            // matches Excel's output, e.g. 2026-03-19T04:24:56.77
            return value.ToString("yyyy-MM-ddTHH:mm:ss.fffffff", CultureInfo.InvariantCulture)
                .TrimEnd('0')
                .TrimEnd('.');
        }

        /// <summary>
        /// Cell reference, e.g. "B5"
        /// </summary>
        [XmlAttribute("ref")]
        public string @ref
        {
            get { return this.refField; }
            set { this.refField = value; }
        }

        /// <summary>
        /// Comment creation timestamp
        /// </summary>
        [XmlAttribute("dT")]
        public DateTime? dT
        {
            get { return this.dTField; }
            set { this.dTField = value; }
        }

        /// <summary>
        /// The id of the person in the persons part that authored this comment
        /// </summary>
        [XmlAttribute("personId")]
        public string personId
        {
            get { return this.personIdField; }
            set { this.personIdField = value; }
        }

        /// <summary>
        /// Unique id of this comment
        /// </summary>
        [XmlAttribute("id")]
        public string id
        {
            get { return this.idField; }
            set { this.idField = value; }
        }

        /// <summary>
        /// The id of the comment this one is a reply to, null for top level comments
        /// </summary>
        [XmlAttribute("parentId")]
        public string parentId
        {
            get { return this.parentIdField; }
            set { this.parentIdField = value; }
        }

        [XmlElement("text")]
        public string text
        {
            get { return this.textField; }
            set { this.textField = value; }
        }
    }
}
