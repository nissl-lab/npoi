using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(ElementName = "sst", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public class CT_Sst
    {

        private List<CT_Rst> siField;

        private CT_ExtensionList extLstField;

        private int countField;

        private bool countFieldSpecified;

        private int uniqueCountField;

        private bool uniqueCountFieldSpecified;

        public CT_Sst()
        {
            //this.extLstField = new CT_ExtensionList();
            this.siField = new List<CT_Rst>();
        }

        [XmlElement]
        public List<CT_Rst> si
        {
            get
            {
                return this.siField;
            }
            set
            {
                this.siField = value;
            }
        }
        [XmlElement]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }
        [XmlAttribute]
        public int count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
                this.countFieldSpecified = true;
            }
        }
        public bool countSpecified
        {
            get { return this.countFieldSpecified; }
            set { this.countFieldSpecified = value; }
        }
        [XmlAttribute]
        public int uniqueCount
        {
            get
            {
                return this.uniqueCountField;
            }
            set
            {
                this.uniqueCountField = value;
                this.uniqueCountFieldSpecified = true;
            }
        }
        public bool uniqueCountSpecified
        {
            get { return this.uniqueCountFieldSpecified; }
            set { this.uniqueCountFieldSpecified = value; }
        }

    }
}
