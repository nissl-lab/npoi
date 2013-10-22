using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Borders
    {
        private List<CT_Border> borderField;
        private uint countField = 0;
        private bool countFieldSpecified = false;

        public CT_Borders()
        {
            this.borderField = new List<CT_Border>();
        }
        public CT_Border AddNewBorder()
        {
            CT_Border border = new CT_Border();
            this.borderField.Add(border);
            return border;
        }
        [XmlElement]
        public List<CT_Border> border
        {
            get
            {
                return this.borderField;
            }
            set
            {
                this.borderField = value;
            }
        }
        public void SetBorderArray(List<CT_Border> array)
        {
            borderField = array;
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
