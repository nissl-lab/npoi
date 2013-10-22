using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_NumFmt
    {

        private uint numFmtIdField;

        private string formatCodeField;

        [XmlAttribute]
        public uint numFmtId
        {
            get
            {
                return this.numFmtIdField;
            }
            set
            {
                this.numFmtIdField = value;
            }
        }
        [XmlAttribute]
        public string formatCode
        {
            get
            {
                return this.formatCodeField;
            }
            set
            {
                this.formatCodeField = value;
            }
        }
    }
}
