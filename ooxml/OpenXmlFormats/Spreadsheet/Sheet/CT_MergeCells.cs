using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;
using System.Xml;
using NPOI.OpenXml4Net.Util;
using System.IO;


namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_MergeCells
    {

        private List<CT_MergeCell> mergeCellField;

        private uint countField;

        private bool countFieldSpecified;

        public static CT_MergeCells Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MergeCells ctObj = new CT_MergeCells();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes["count"]);
            ctObj.mergeCell = new List<CT_MergeCell>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "mergeCell")
                    ctObj.mergeCell.Add(CT_MergeCell.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "count", this.count);
            sw.Write(">");
            if (this.mergeCell != null)
            {
                foreach (CT_MergeCell x in this.mergeCell)
                {
                    if (x != null) {
                        x.Write(sw, "mergeCell");
                    }
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }


        public CT_MergeCells()
        {
            this.mergeCellField = new List<CT_MergeCell>();
        }
        public CT_MergeCell GetMergeCellArray(int index)
        {
            return this.mergeCellField[index];
        }
        public void SetMergeCellArray(CT_MergeCell[] array)
        {
            mergeCell = new List<CT_MergeCell>(array);
        }
        public int sizeOfMergeCellArray()
        {
            return mergeCell.Count;
        }
        public CT_MergeCell AddNewMergeCell()
        {
            CT_MergeCell mergecell = new CT_MergeCell();
            mergeCell.Add(mergecell);
            return mergecell;
        }
        [XmlElement]
        public List<CT_MergeCell> mergeCell
        {
            get
            {
                return this.mergeCellField;
            }
            set
            {
                this.mergeCellField = value;
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
