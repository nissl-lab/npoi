using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Cols
    {

        private List<CT_Col> colField = new List<CT_Col>(); // required

        //public CT_Cols()
        //{
        //    this.colField = new List<CT_Col>();
        //}
        public void SetColArray(List<CT_Col> array)
        {
            colField = array;
        }
        public CT_Col AddNewCol()
        {
            CT_Col newCol = new CT_Col();
            this.colField.Add(newCol);
            return newCol;
        }
        public CT_Col InsertNewCol(int index)
        {
            CT_Col newCol = new CT_Col();
            this.colField.Insert(index, newCol);
            return newCol;
        }
        public void RemoveCol(int index)
        {
            this.colField.RemoveAt(index);
        }

        public int sizeOfColArray()
        {
            return col.Count;
        }
        public CT_Col GetColArray(int index)
        {
            return colField[index];
        }


        public List<CT_Col> GetColArray()
        {
            return colField;
        }
        [XmlElement]
        public List<CT_Col> col
        {
            get
            {
                return this.colField;
            }
            set
            {
                this.colField = value;
            }
        }
    }
}
