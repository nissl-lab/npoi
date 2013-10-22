using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    /// <summary>
    /// Holds the Column Width and its Formatting
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Col
    {

        private uint minField; // required

        private uint maxField; // required

        private double widthField;  // optional, but width has no default value
        private bool widthSpecifiedField;

        private uint styleField;// optional, as are the following attributes
        private bool styleSpecifiedField;

        private bool hiddenField;

        private bool bestFitField;

        private bool customWidthField;

        private bool phoneticField;

        private byte outlineLevelField;

        private bool collapsedField = true;
        private bool collapsedSpecifiedField = true;

        [XmlAttribute]
        public uint min
        {
            get
            {
                return this.minField;
            }
            set
            {
                this.minField = value;
            }
        }

        [XmlAttribute]
        public uint max
        {
            get
            {
                return this.maxField;
            }
            set
            {
                this.maxField = value;
            }
        }

        [XmlAttribute]
        public double width
        {
            get
            {
                return this.widthField;
            }
            set
            {
                this.widthField = value;
                this.widthSpecified = true;
            }
        }
        [XmlIgnore]
        public bool widthSpecified
        {
            get
            {
                return this.widthSpecifiedField;
            }
            set
            {
                this.widthSpecifiedField = value;
            }
        }


        public bool IsSetBestFit()
        {
            return this.bestFitField != false;
        }
        public bool IsSetCustomWidth()
        {
            return this.customWidthField != false;
        }
        public bool IsSetHidden()
        {
            return this.hiddenField != false;
        }
        public bool IsSetStyle()
        {
            return this.styleSpecifiedField;
        }
        public bool IsSetWidth()
        {
            return this.widthSpecifiedField;
        }
        public bool IsSetCollapsed()
        {
            return this.collapsedSpecifiedField;
        }
        public bool IsSetPhonetic()
        {
            return this.phoneticField != false;
        }
        public bool IsSetOutlineLevel()
        {
            return this.outlineLevelField != 0;
        }
        public void UnsetHidden()
        {
            this.hiddenField = false;
        }
        public void UnsetCollapsed()
        {
            this.collapsedField = true;
            this.collapsedSpecified = false;
        }



        [DefaultValue(typeof(uint), "0")]
        [XmlAttribute]
        public uint style
        {
            get
            {
                return styleField;
            }
            set
            {
                this.styleField = value;
            }
        }
        [XmlIgnore]
        public bool styleSpecified
        {
            get { return this.styleSpecifiedField; }
            set { this.styleSpecifiedField = value; }
        }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool hidden
        {
            get
            {
                return hiddenField;
            }
            set
            {
                this.hiddenField = value;
            }
        }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool bestFit
        {
            get
            {
                return bestFitField;
            }
            set
            {
                this.bestFitField = value;
            }
        }
        [DefaultValue(false)]
        [XmlAttribute]
        public bool customWidth
        {
            get
            {
                return customWidthField;
            }
            set
            {
                this.customWidthField = value;
            }
        }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool phonetic
        {
            get
            {
                return phoneticField;
            }
            set
            {
                this.phoneticField = value;
            }
        }


        [DefaultValue(typeof(byte), "0")]
        [XmlAttribute]
        public byte outlineLevel
        {
            get
            {
                return outlineLevelField;
            }
            set
            {
                this.outlineLevelField = value;
            }
        }

        [DefaultValue(true)]
        [XmlAttribute]
        public bool collapsed
        {
            get
            {
                return collapsedField;
            }
            set
            {
                this.collapsedField = value;
                this.collapsedSpecifiedField = true;
            }
        }
        [XmlIgnore]
        public bool collapsedSpecified
        {
            get { return this.collapsedSpecifiedField; }
            set { this.collapsedSpecifiedField = value; }
        }


    }
}
