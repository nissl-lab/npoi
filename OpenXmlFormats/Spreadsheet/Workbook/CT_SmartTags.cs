using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_SmartTagTypes
    {

        private List<CT_SmartTagType> smartTagTypeField;

        public CT_SmartTagTypes()
        {
            //this.smartTagTypeField = new List<CT_SmartTagType>();
        }
        public static CT_SmartTagTypes Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SmartTagTypes ctObj = new CT_SmartTagTypes();
            ctObj.smartTagType = new List<CT_SmartTagType>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "smartTagType")
                    ctObj.smartTagType.Add(CT_SmartTagType.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            if (this.smartTagType != null)
            {
                foreach (CT_SmartTagType x in this.smartTagType)
                {
                    x.Write(sw, "smartTagType");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlElement]
        public List<CT_SmartTagType> smartTagType
        {
            get
            {
                return this.smartTagTypeField;
            }
            set
            {
                this.smartTagTypeField = value;
            }
        }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_SmartTagType
    {
        public static CT_SmartTagType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SmartTagType ctObj = new CT_SmartTagType();
            ctObj.namespaceUri = XmlHelper.ReadString(node.Attributes["namespaceUri"]);
            ctObj.name = XmlHelper.ReadString(node.Attributes["name"]);
            ctObj.url = XmlHelper.ReadString(node.Attributes["url"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "namespaceUri", this.namespaceUri);
            XmlHelper.WriteAttribute(sw, "name", this.name);
            XmlHelper.WriteAttribute(sw, "url", this.url);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        private string namespaceUriField;

        private string nameField;

        private string urlField;

        public string namespaceUri
        {
            get
            {
                return this.namespaceUriField;
            }
            set
            {
                this.namespaceUriField = value;
            }
        }

        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        public string url
        {
            get
            {
                return this.urlField;
            }
            set
            {
                this.urlField = value;
            }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_SmartTagPr
    {

        private bool embedField;

        private ST_SmartTagShow showField;
        public static CT_SmartTagPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SmartTagPr ctObj = new CT_SmartTagPr();
            ctObj.embed = XmlHelper.ReadBool(node.Attributes["embed"]);
            if (node.Attributes["show"] != null)
                ctObj.show = (ST_SmartTagShow)Enum.Parse(typeof(ST_SmartTagShow), node.Attributes["show"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "embed", this.embed);
            XmlHelper.WriteAttribute(sw, "show", this.show.ToString());
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_SmartTagPr()
        {
            this.embedField = false;
            this.showField = ST_SmartTagShow.all;
        }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool embed
        {
            get
            {
                return this.embedField;
            }
            set
            {
                this.embedField = value;
            }
        }
        [XmlAttribute]
        [DefaultValue(ST_SmartTagShow.all)]
        public ST_SmartTagShow show
        {
            get
            {
                return this.showField;
            }
            set
            {
                this.showField = value;
            }
        }
    }
}
