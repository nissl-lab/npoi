using System;
using System.Collections.Generic;
using System.Linq;
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
                if (childNode.LocalName == nameof(row))
                    ctObj.row.Add(CT_Row.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            sw.Write(">");
            this.row?.ForEach(x => x.Write(sw, nameof(row)));
            sw.Write($"</{nodeName}>");
        }

        //public CT_SheetData()
        //{
        //    this.rowField = new List<CT_Row>();
        //}
        public CT_Row AddNewRow()
        {
            if (null == row) { row = new List<CT_Row>(); }
            CT_Row newrow = new CT_Row();
            row.Add(newrow);
            return newrow;
        }
        public CT_Row InsertNewRow(int index)
        {
            if (null == row) { row = new List<CT_Row>(); }
            CT_Row newrow = new CT_Row();
            row.Insert(index, newrow);
            return newrow;
        }
        public void RemoveRows(IList<CT_Row> toRemove)
        {
            if (row == null) return;
            foreach (CT_Row r in toRemove)
            {
                row.Remove(r);
            }
        }
        public void RemoveRow(int rowNum)
        {
            if (null != row)
            {
                CT_Row rowToRemove=null;
                foreach (CT_Row ctrow in row)
                {
                    if (ctrow.r == rowNum)
                    {
                        rowToRemove = ctrow;
                        break;
                    }
                }
                row.Remove(rowToRemove);
            }
        }
        public int SizeOfRowArray()
        {
            return row?.Count ?? 0;
        }

        public CT_Row GetRowArray(int index)
        {
            return row?[index];
        }
        [XmlElement(nameof(row))]
        public List<CT_Row> row { get; set; } = null;
        [XmlIgnore]
        public bool rowSpecified
        {
            get { return null != row; }
        }
    }
}
