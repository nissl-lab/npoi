using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Rst
    {

        private string tField = null; // optional field -> initialize as null so that it is not serialized by default.

        private List<CT_RElt> rField = null; // optional field 

        private List<CT_PhoneticRun> rPhField = null; // optional field 

        private CT_PhoneticPr phoneticPrField = null; // optional field 

        public void Set(CT_Rst o)
        {
            this.tField = o.tField;
            this.rField = o.rField;
            this.rPhField = o.rPhField;
            this.phoneticPrField = o.phoneticPrField;
        }

        #region t
        public bool IsSetT()
        {
            return this.tField != null;
        }
        public void unsetT()
        {
            this.tField = null;
        }
        [XmlElement("t", DataType = "string")]
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
        #endregion t

        #region r
        /// <summary>
        /// Rich Text Run
        /// </summary>
        [XmlElement("r")]
        public List<CT_RElt> r
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
        private string xmltext;
        [XmlText]
        public string XmlText
        {
            get { return xmltext; }
            set { xmltext = value; }
        }
        public CT_RElt AddNewR()
        {
            if (null == rField) { rField = new List<CT_RElt>(); }
            CT_RElt r = new CT_RElt();
            this.rField.Add(r);
            return r;
        }
        public int sizeOfRArray()
        {
            return (null == rField) ? 0 : r.Count;
        }
        public CT_RElt GetRArray(int index)
        {
            return (null == rField) ? null : this.rField[index];
        }
        #endregion r

        /// <summary>
        /// Phonetic Run
        /// </summary>
        [XmlElement("rPh")]
        public List<CT_PhoneticRun> rPh
        {
            get
            {
                return this.rPhField;
            }
            set
            {
                this.rPhField = value;
            }
        }
        /// <summary>
        /// Phonetic Properties
        /// </summary>
        [XmlElement("phoneticPr")]
        public CT_PhoneticPr phoneticPr
        {
            get
            {
                return this.phoneticPrField;
            }
            set
            {
                this.phoneticPrField = value;
            }
        }

    }
}
