using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_MergeCells
    {

        private List<CT_MergeCell> mergeCellField;

        private uint countField;

        private bool countFieldSpecified;



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
