using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_CellStyles
    {

        private List<CT_CellStyle> cellStyleField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_CellStyles()
        {
            //this.cellStyleField = new List<CT_CellStyle>();
        }
        [XmlElement]
        public List<CT_CellStyle> cellStyle
        {
            get
            {
                return this.cellStyleField;
            }
            set
            {
                this.cellStyleField = value;
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
