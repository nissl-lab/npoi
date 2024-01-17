using NPOI.SS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
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
            if (null == colField)
            {
                colField = new List<CT_Col>();
            }

            CT_Col newCol = new CT_Col();
            this.colField.Add(newCol);
            return newCol;
        }

        public CT_Col InsertNewCol(int index)
        {
            if (null == colField)
            {
                colField = new List<CT_Col>();
            }

            CT_Col newCol = new CT_Col();
            this.colField.Insert(index, newCol);
            return newCol;
        }
        public void RemoveCol(int index)
        {
            this.colField.RemoveAt(index);
        }
        public void RemoveCols(IList<CT_Col> toRemove)
        {
            if (colField == null)
            {
                return;
            }

            foreach (CT_Col c in toRemove)
            {
                _ = colField.Remove(c);
            }
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

        public static CT_Cols Parse(XmlNode node, XmlNamespaceManager namespaceManager, int lastColumn)
        {
            if (node == null)
            {
                return null;
            }

            CT_Cols ctObj = new CT_Cols
            {
                col = new List<CT_Col>()
            };
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "col")
                {
                    CT_Col ctCol = CT_Col.Parse(childNode, namespaceManager);

                    if (ctCol.min != ctCol.max)
                    {
                        BreakUpCtCol(ctObj, ctCol, lastColumn);
                    }
                    else
                    {
                        ctObj.col.Add(ctCol);
                    }
                }
            }

            return ctObj;
        }

        /// <summary>
        /// For ease of use of columns in NPOI break up <see cref="CT_Col"/>s
        /// that span over multiple physical columns into individual
        /// <see cref="CT_Col"/>s for each physical column.
        /// </summary>
        /// <param name="ctObj"></param>
        /// <param name="ctCol"></param>
        private static void BreakUpCtCol(CT_Cols ctObj, CT_Col ctCol, int lastColumn)
        {
            int max = ctCol.max >= SpreadsheetVersion.EXCEL2007.LastColumnIndex - 1
                ? lastColumn
                : (int)ctCol.max;

            for (int i = (int)ctCol.min; i <= max; i++)
            {
                CT_Col breakOffCtCol = ctCol.Copy();
                breakOffCtCol.min = (uint)i;
                breakOffCtCol.max = (uint)i;

                ctObj.col.Add(breakOffCtCol);
            }
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            List<CT_Col> combinedCols = CombineCols(col);

            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");

            if (combinedCols != null)
            {
                foreach (CT_Col x in combinedCols)
                {
                    x.Write(sw, "col");
                }
            }

            sw.Write(string.Format("</{0}>", nodeName));
        }

        /// <summary>
        /// Broken up by the <see cref="BreakUpCtCol(CT_Cols, CT_Col)"/> method
        /// <see cref="CT_Col"/>s are combined into <see cref="CT_Col"/> spans
        /// </summary>
        private static List<CT_Col> CombineCols(List<CT_Col> cols)
        {
            List<CT_Col> combinedCols = new List<CT_Col>();

            cols.Sort((c1, c2) => c1.min.CompareTo(c2.min));

            CT_Col lastCol = null;

            foreach (CT_Col col in cols)
            {
                if (lastCol == null)
                {
                    lastCol = col;
                    continue;
                }

                if (col.IsAdjacentAndCanBeCombined(lastCol))
                {
                    lastCol.CombineWith(col);
                    continue;
                }

                combinedCols.Add(lastCol);
                lastCol = col;
            }

            if (lastCol != null)
            {
                combinedCols.Add(lastCol);
            }

            return combinedCols;
        }

        public void SetColArray(CT_Col[] colArray)
        {
            this.colField = new List<CT_Col>(colArray);
        }
    }
}
