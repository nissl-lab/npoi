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
    public class CT_CellStyles
    {

        private List<CT_CellStyle> cellStyleField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_CellStyles()
        {
            //this.cellStyleField = new List<CT_CellStyle>();
        }
        [XmlElement]
        public List<CT_CellStyle> cellStyle
        {
            get
            {
                return this.cellStyleField;
            }
            set
            {
                this.cellStyleField = value;
            }
        }

        [XmlAttribute]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [XmlIgnore]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }

        public static CT_CellStyles Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CellStyles ctObj = new CT_CellStyles();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes["count"]);
            ctObj.cellStyle = new List<CT_CellStyle>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "cellStyle")
                    ctObj.cellStyle.Add(CT_CellStyle.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "count", this.count);
            sw.Write(">");
            if (this.cellStyle != null)
            {
                foreach (CT_CellStyle x in this.cellStyle)
                {
                    x.Write(sw, "cellStyle");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

    }
}
