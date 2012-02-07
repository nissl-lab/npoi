using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class CT_Color
    {
        private bool autoField;

        private bool autoFieldSpecified;

        private long indexedField;

        private bool indexedFieldSpecified;

        private byte[] rgbField;

        private long themeField;

        private bool themeFieldSpecified;

        private double tintField;

        public CT_Color()
        {
            this.tintField = 0D;
        }

        public bool auto
        {
            get
            {
                return this.autoField;
            }
            set
            {
                this.autoField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool autoSpecified
        {
            get
            {
                return this.autoFieldSpecified;
            }
            set
            {
                this.autoFieldSpecified = value;
            }
        }

        public long indexed
        {
            get
            {
                return this.indexedField;
            }
            set
            {
                this.indexedField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool indexedSpecified
        {
            get
            {
                return this.indexedFieldSpecified;
            }
            set
            {
                this.indexedFieldSpecified = value;
            }
        }

        public byte[] rgb
        {
            get
            {
                return this.rgbField;
            }
            set
            {
                this.rgbField = value;
            }
        }

        public long theme
        {
            get
            {
                return this.themeField;
            }
            set
            {
                this.themeField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool themeSpecified
        {
            get
            {
                return this.themeFieldSpecified;
            }
            set
            {
                this.themeFieldSpecified = value;
            }
        }

        [System.ComponentModel.DefaultValueAttribute(0D)]
        public double tint
        {
            get
            {
                return this.tintField;
            }
            set
            {
                this.tintField = value;
            }
        }

        public void SetRgb(byte R, byte G, byte B)
        {
            
        }
        public bool IsSetRgb()
        {
            return rgb != null;
        }
        public void SetRgb(byte[] rgb)
        {
            Array.Copy(rgb, this.rgb, 3);
        }
        public byte[] GetRgb()
        {
            return rgb;
        }

    }
}
