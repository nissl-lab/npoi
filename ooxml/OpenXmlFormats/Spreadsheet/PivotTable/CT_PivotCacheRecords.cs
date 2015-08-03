using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("pivotCacheRecords", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public partial class CT_PivotCacheRecords
    {

        private List<object> rField;

        private CT_ExtensionList extLstField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_PivotCacheRecords()
        {
            this.extLstField = new CT_ExtensionList();
            this.rField = new List<object>();
        }

        [System.Xml.Serialization.XmlArrayAttribute(Order = 0)]
        [System.Xml.Serialization.XmlArrayItemAttribute("b", typeof(CT_Boolean), IsNullable = false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("d", typeof(CT_DateTime), IsNullable = false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("e", typeof(CT_Error), IsNullable = false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("m", typeof(CT_Missing), IsNullable = false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("n", typeof(CT_Number), IsNullable = false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("s", typeof(CT_String), IsNullable = false)]
        [System.Xml.Serialization.XmlArrayItemAttribute("x", typeof(CT_Index), IsNullable = false)]
        public List<object> r
        {
            get
            {
                return this.rField;
            }
            set
            {
                this.rField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        public static CT_PivotCacheRecords Parse(System.IO.Stream is1)
        {
            throw new NotImplementedException();
        }

        public void Save(System.IO.Stream out1)
        {
            throw new NotImplementedException();
        }
    }
    
}
