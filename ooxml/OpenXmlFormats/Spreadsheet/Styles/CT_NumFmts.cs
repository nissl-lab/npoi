using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_NumFmts
    {

        private List<CT_NumFmt> numFmtField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_NumFmts()
        {
            this.numFmtField = new List<CT_NumFmt>();
        }

        public CT_NumFmt AddNewNumFmt()
        {
            CT_NumFmt newNumFmt = new CT_NumFmt();
            this.numFmtField.Add(newNumFmt);
            return newNumFmt;
        }
        [XmlElement]
        public List<CT_NumFmt> numFmt
        {
            get
            {
                return this.numFmtField;
            }
            set
            {
                this.numFmtField = value;
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
