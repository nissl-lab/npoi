using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_SheetData
    {

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
