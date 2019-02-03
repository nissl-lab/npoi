using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;
using System.Xml;
using System.IO;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_SheetData
    {

        public static CT_SheetData Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SheetData ctObj = new CT_SheetData();
            ctObj.row = new List<CT_Row>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "row")
                    ctObj.row.Add(CT_Row.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            if (this.row != null)
            {
                foreach (CT_Row x in this.row)
                {
                    x.Write(sw, "row");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }



        private List<CT_Row> rowField = null; // [0..*] 

        //public CT_SheetData()
        //{
        //    this.rowField = new List<CT_Row>();
        //}
        public CT_Row AddNewRow()
        {
            if (null == rowField) { rowField = new List<CT_Row>(); }
            CT_Row newrow = new CT_Row();
            rowField.Add(newrow);
            return newrow;
        }
        public CT_Row InsertNewRow(int index)
        {
            if (null == rowField) { rowField = new List<CT_Row>(); }
            CT_Row newrow = new CT_Row();
            rowField.Insert(index, newrow);
            return newrow;
        }
        public void RemoveRows(IList<CT_Row> toRemove)
        {
            if (rowField == null) return;
            foreach (CT_Row r in toRemove)
            {
                rowField.Remove(r);
            }
        }
        public void RemoveRow(int rowNum)
        {
            if (null != rowField)
            {
                CT_Row rowToRemove=null;
                foreach (CT_Row ctrow in rowField)
                {
                    if (ctrow.r == rowNum)
                    {
                        rowToRemove = ctrow;
                        break;
                    }
                }
                rowField.Remove(rowToRemove);
            }
        }
        public int SizeOfRowArray()
        {
            return (null == rowField) ? 0 : rowField.Count;
        }

        public CT_Row GetRowArray(int index)
        {
            return (null == rowField) ? null : rowField[index];
        }
        [XmlElement("row")]
        public List<CT_Row> row
        {
            get
            {
                return this.rowField;
            }
            set
            {
                this.rowField = value;
            }
        }
        [XmlIgnore]
        public bool rowSpecified
        {
            get { return null != rowField; }
        }
    }
}
