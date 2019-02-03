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
    public class CT_CellStyleXfs
    {

        private List<CT_Xf> xfField = null;

        private uint countField;

        private bool countFieldSpecified;
        public static CT_CellStyleXfs Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CellStyleXfs ctObj = new CT_CellStyleXfs();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes["count"]);
            ctObj.xf = new List<CT_Xf>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "xf")
                    ctObj.xf.Add(CT_Xf.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "count", this.count);
            sw.Write(">");
            if (this.xf != null)
            {
                foreach (CT_Xf x in this.xf)
                {
                    x.Write(sw, "xf");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_CellStyleXfs()
        {
            //this.xfField = new List<CT_Xf>();
        }
        public CT_Xf AddNewXf()
        {
            if (this.xfField == null)
                this.xfField = new List<CT_Xf>();
            CT_Xf xf = new CT_Xf();
            this.xfField.Add(xf);
            return xf;
        }
        [XmlElement]
        public List<CT_Xf> xf
        {
            get
            {
                return this.xfField;
            }
            set
            {
                this.xfField = value;
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
    }
}
