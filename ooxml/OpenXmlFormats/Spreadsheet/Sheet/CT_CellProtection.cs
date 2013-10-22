using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_CellProtection
    {

        private bool? lockedField = null;

        private bool? hiddenField = null;

        public bool IsSetHidden()
        {
            return this.hiddenField != null;
        }
        public bool IsSetLocked()
        {
            return this.lockedField != null;
        }

        [XmlAttribute]
        public bool locked
        {
            get
            {
                return (bool)this.lockedField;
            }
            set
            {
                this.lockedField = value;
            }
        }
        [XmlIgnore]
        public bool lockedSpecified
        {
            get { return null != this.lockedField; }
        }

        [XmlAttribute]
        public bool hidden
        {
            get
            {
                return (bool)this.hiddenField;
            }
            set
            {
                this.hiddenField = value;
            }
        }
        [XmlIgnore]
        public bool hiddenSpecified
        {
            get { return null != this.hiddenField; }
        }
    }
}
