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
    public class CT_Fills
    {
        private List<CT_Fill> fillField;
        private uint countField = 0;
        private bool countFieldSpecified = false;

        public CT_Fills()
        {
            //this.fillField = new List<CT_Fill>();
        }
        public static CT_Fills Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Fills ctObj = new CT_Fills();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes["count"]);
            ctObj.fill = new List<CT_Fill>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "fill")
                    ctObj.fill.Add(CT_Fill.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "count", this.count);
            sw.Write(">");
            if (this.fill != null)
            {
                foreach (CT_Fill x in this.fill)
                {
                    x.Write(sw, "fill");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlElement]
        public List<CT_Fill> fill
        {
            get
            {
                return this.fillField;
            }
            set
            {
                this.fillField = value;
            }
        }
        public void SetFillArray(List<CT_Fill> array)
        {
            fillField = array;
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
