using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public enum ST_HorizontalAlignment
    {
        general,
        left,
        center,
        right,
        fill,
        justify,
        centerContinuous,
        distributed,
    }

    public enum ST_VerticalAlignment
    {
        top,
        center,
        bottom,
        justify,
        distributed,
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_CellAlignment
    {

        private ST_HorizontalAlignment horizontalField;

        private bool horizontalFieldSpecified;

        private ST_VerticalAlignment verticalField;

        private bool verticalFieldSpecified;

        private long textRotationField;

        private bool textRotationFieldSpecified;

        private bool wrapTextField;

        private bool wrapTextFieldSpecified;

        private long indentField;

        private bool indentFieldSpecified;

        private int relativeIndentField;

        private bool relativeIndentFieldSpecified;

        private bool justifyLastLineField;

        private bool justifyLastLineFieldSpecified;

        private bool shrinkToFitField;

        private bool shrinkToFitFieldSpecified;

        private long readingOrderField;

        private bool readingOrderFieldSpecified;

        public bool IsSetHorizontal()
        {
            return this.horizontalField != ST_HorizontalAlignment.general;
        }
        public bool IsSetVertical()
        {
            return this.verticalField != ST_VerticalAlignment.top;
        }

        [XmlAttribute]
        [DefaultValue(ST_HorizontalAlignment.general)]
        public ST_HorizontalAlignment horizontal
        {
            get
            {
                return this.horizontalField;
            }
            set
            {
                this.horizontalField = value;
            }
        }
        [XmlIgnore]
        public bool horizontalSpecified
        {
            get
            {
                return this.horizontalFieldSpecified;
            }
            set
            {
                this.horizontalFieldSpecified = value;
            }
        }
        [XmlAttribute]
        [DefaultValue(ST_VerticalAlignment.top)]
        public ST_VerticalAlignment vertical
        {
            get
            {
                return this.verticalField;
            }
            set
            {
                this.verticalField = value;
            }
        }

        [XmlIgnore]
        public bool verticalSpecified
        {
            get
            {
                return this.verticalFieldSpecified;
            }
            set
            {
                this.verticalFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public long textRotation
        {
            get
            {
                return this.textRotationField;
            }
            set
            {
                this.textRotationField = value;
            }
        }

        [XmlIgnore]
        public bool textRotationSpecified
        {
            get
            {
                return this.textRotationFieldSpecified;
            }
            set
            {
                this.textRotationFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public bool wrapText
        {
            get
            {
                return this.wrapTextField;
            }
            set
            {
                this.wrapTextField = value;
            }
        }

        [XmlIgnore]
        public bool wrapTextSpecified
        {
            get
            {
                return this.wrapTextFieldSpecified;
            }
            set
            {
                this.wrapTextFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public long indent
        {
            get
            {
                return this.indentField;
            }
            set
            {
                this.indentField = value;
            }
        }

        [XmlIgnore]
        public bool indentSpecified
        {
            get
            {
                return this.indentFieldSpecified;
            }
            set
            {
                this.indentFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public int relativeIndent
        {
            get
            {
                return this.relativeIndentField;
            }
            set
            {
                this.relativeIndentField = value;
            }
        }

        [XmlIgnore]
        public bool relativeIndentSpecified
        {
            get
            {
                return this.relativeIndentFieldSpecified;
            }
            set
            {
                this.relativeIndentFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public bool justifyLastLine
        {
            get
            {
                return this.justifyLastLineField;
            }
            set
            {
                this.justifyLastLineField = value;
            }
        }

        [XmlIgnore]
        public bool justifyLastLineSpecified
        {
            get
            {
                return this.justifyLastLineFieldSpecified;
            }
            set
            {
                this.justifyLastLineFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public bool shrinkToFit
        {
            get
            {
                return this.shrinkToFitField;
            }
            set
            {
                this.shrinkToFitField = value;
            }
        }

        [XmlIgnore]
        public bool shrinkToFitSpecified
        {
            get
            {
                return this.shrinkToFitFieldSpecified;
            }
            set
            {
                this.shrinkToFitFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public long readingOrder
        {
            get
            {
                return this.readingOrderField;
            }
            set
            {
                this.readingOrderField = value;
            }
        }

        [XmlIgnore]
        public bool readingOrderSpecified
        {
            get
            {
                return this.readingOrderFieldSpecified;
            }
            set
            {
                this.readingOrderFieldSpecified = value;
            }
        }
    }
}
