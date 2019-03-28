using System;

using System.Xml.Serialization;
using System.Xml;
using NPOI.OpenXml4Net.Util;
using System.IO;
using NPOI.OpenXml4Net.OPC;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Hyperlink
    {

        private string refField = null;

        private string idField = null; // this and the other ones are optional

        private string locationField = null;

        private string tooltipField = null;

        private string displayField = null;

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

        [XmlAttribute(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }

        [XmlAttribute]
        public string location
        {
            get
            {
                return this.locationField;
            }
            set
            {
                this.locationField = value;
            }
        }

        [XmlAttribute]
        public string tooltip
        {
            get
            {
                return this.tooltipField;
            }
            set
            {
                this.tooltipField = value;
            }
        }

        [XmlAttribute]
        public string display
        {
            get
            {
                return this.displayField;
            }
            set
            {
                this.displayField = value;
            }
        }

        public static CT_Hyperlink Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Hyperlink ctObj = new CT_Hyperlink();
            ctObj.@ref = XmlHelper.ReadString(node.Attributes["ref"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["id", PackageNamespaces.SCHEMA_RELATIONSHIPS]);
            ctObj.location = XmlHelper.ReadString(node.Attributes["location"]);
            ctObj.tooltip = XmlHelper.ReadString(node.Attributes["tooltip"]);
            ctObj.display = XmlHelper.ReadString(node.Attributes["display"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "ref", this.@ref);
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            XmlHelper.WriteAttribute(sw, "location", this.location);
            XmlHelper.WriteAttribute(sw, "tooltip", this.tooltip);
            XmlHelper.WriteAttribute(sw, "display", this.display);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }







    }
}
