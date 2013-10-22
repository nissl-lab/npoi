using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    /// <summary>
    /// Rich Text Run container.
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_RElt
    {

        private CT_RPrElt rPrField = null; // optional field 

        private string tField = string.Empty; // required field 

        public CT_RPrElt AddNewRPr()
        {
            this.rPrField = new CT_RPrElt();
            return rPrField;
        }

        /// <summary>
        /// Run Properties
        /// </summary>
        [XmlElement("rPr")]
        public CT_RPrElt rPr
        {
            get
            {
                return this.rPrField;
            }
            set
            {
                this.rPrField = value;
            }
        }

        /// <summary>
        /// Text
        /// </summary>
        [XmlElement("t")]
        public string t
        {
            get
            {
                return this.tField;
            }
            set
            {
                this.tField = value;
            }
        }
    }
}
