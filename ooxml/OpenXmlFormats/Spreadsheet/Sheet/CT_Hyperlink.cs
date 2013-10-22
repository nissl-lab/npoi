using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Hyperlink
    {

        private string refField = string.Empty; // Required

        private string idField = null; // this and the other ones are optional

        private string locationField = null;

        private string tooltipField = null;

        private string displayField = null;

        [XmlAttribute("ref")]
        public string @ref
        {
            get
            {
                return this.refField;
            }
            set
            {
                this.refField = value;
            }
        }

        [XmlAttribute(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }

        [XmlAttribute]
        public string location
        {
            get
            {
                return this.locationField;
            }
            set
            {
                this.locationField = value;
            }
        }

        [XmlAttribute]
        public string tooltip
        {
            get
            {
                return this.tooltipField;
            }
            set
            {
                this.tooltipField = value;
            }
        }

        [XmlAttribute]
        public string display
        {
            get
            {
                return this.displayField;
            }
            set
            {
                this.displayField = value;
            }
        }
    }
}
