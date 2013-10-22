using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_CellStyleXfs
    {

        private List<CT_Xf> xfField = null;

        private uint countField;

        private bool countFieldSpecified;

        public CT_CellStyleXfs()
        {
            //this.xfField = new List<CT_Xf>();
        }
        public CT_Xf AddNewXf()
        {
            if (this.xfField == null)
                this.xfField = new List<CT_Xf>();
            CT_Xf xf = new CT_Xf();
            this.xfField.Add(xf);
            return xf;
        }
        [XmlElement]
        public List<CT_Xf> xf
        {
            get
            {
                return this.xfField;
            }
            set
            {
                this.xfField = value;
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
