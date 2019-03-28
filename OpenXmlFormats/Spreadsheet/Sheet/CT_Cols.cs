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


        public List<CT_Col> GetColList()
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

        public static CT_Cols Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Cols ctObj = new CT_Cols();
            ctObj.col = new List<CT_Col>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "col")
                    ctObj.col.Add(CT_Col.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            if (this.col != null)
            {
                foreach (CT_Col x in this.col)
                {
                    x.Write(sw, "col");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }



        public void SetColArray(CT_Col[] colArray)
        {
            this.colField = new List<CT_Col>(colArray);
        }
    }
}
