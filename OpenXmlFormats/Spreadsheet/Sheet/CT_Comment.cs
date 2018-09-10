using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ElementName = "comment")]
    public class CT_Comment
    {

        private CT_Rst textField = new CT_Rst(); // required element 

        private string refField = string.Empty; // required attribute

        private uint authorIdField = 0; // required attribute

        private string guidField = null; // optional attribute

        //public CT_Comment()
        //{
        //    this.textField = new CT_Rst();
        //}
        public static CT_Comment Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Comment ctObj = new CT_Comment();
            ctObj.@ref = XmlHelper.ReadString(node.Attributes["ref"]);
            if (node.Attributes["authorId"] != null)
                ctObj.authorId = XmlHelper.ReadUInt(node.Attributes["authorId"]);
            ctObj.guid = XmlHelper.ReadString(node.Attributes["guid"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "text")
                    ctObj.text = CT_Rst.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }
        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "ref", this.@ref);
            XmlHelper.WriteAttribute(sw, "authorId", this.authorId, true);
            XmlHelper.WriteAttribute(sw, "guid", this.guid);
            sw.Write(">");
            if (this.text != null)
                this.text.Write(sw, "text");
            sw.Write(string.Format("</{0}>", nodeName));
        }
        [XmlElement("text")]
        public CT_Rst text
        {
            get
            {
                return this.textField;
            }
            set
            {
                this.textField = value;
            }
        }

        [XmlAttribute("ref")]
        public string @ref
        {
            get
            {
                return this.refField;
            }
            set
            {
                this.refField = value;
            }
        }

        [XmlAttribute("authorId")]
        public uint authorId
        {
            get
            {
                return this.authorIdField;
            }
            set
            {
                this.authorIdField = value;
            }
        }

        [XmlAttribute("guid")] // 0..1 TODO: Type is ST_Guid
        public string guid
        {
            get
            {
                return this.guidField;
            }
            set
            {
                this.guidField = value;
            }
        }
    }
}
