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
    public class CT_Borders
    {
        private List<CT_Border> borderField;
        private uint countField = 0;
        private bool countFieldSpecified = false;
        public static CT_Borders Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Borders ctObj = new CT_Borders();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes["count"]);
            ctObj.border = new List<CT_Border>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "border")
                    ctObj.border.Add(CT_Border.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "count", this.count, true);
            sw.Write(">");
            if (this.border != null)
            {
                foreach (CT_Border x in this.border)
                {
                    x.Write(sw, "border");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Borders()
        {
            //this.borderField = new List<CT_Border>();
        }
        public CT_Border AddNewBorder()
        {
            if (this.borderField == null)
                this.borderField = new List<CT_Border>();
            CT_Border border = new CT_Border();
            this.borderField.Add(border);
            return border;
        }
        [XmlElement]
        public List<CT_Border> border
        {
            get
            {
                return this.borderField;
            }
            set
            {
                this.borderField = value;
            }
        }
        public void SetBorderArray(List<CT_Border> array)
        {
            borderField = array;
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
    }
}
