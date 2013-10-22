using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Fills
    {
        private List<CT_Fill> fillField;
        private uint countField = 0;
        private bool countFieldSpecified = false;

        public CT_Fills()
        {
            this.fillField = new List<CT_Fill>();
        }
        [XmlElement]
        public List<CT_Fill> fill
        {
            get
            {
                return this.fillField;
            }
            set
            {
                this.fillField = value;
            }
        }
        public void SetFillArray(List<CT_Fill> array)
        {
            fillField = array;
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
