using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml;
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

        private ST_VerticalAlignment verticalField = ST_VerticalAlignment.bottom;

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
        public static CT_CellAlignment Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CellAlignment ctObj = new CT_CellAlignment();
            if (node.Attributes["horizontal"] != null)
                ctObj.horizontal = (ST_HorizontalAlignment)Enum.Parse(typeof(ST_HorizontalAlignment), node.Attributes["horizontal"].Value);
            if (node.Attributes["vertical"] != null)
                ctObj.vertical = (ST_VerticalAlignment)Enum.Parse(typeof(ST_VerticalAlignment), node.Attributes["vertical"].Value);
            ctObj.textRotation = XmlHelper.ReadLong(node.Attributes["textRotation"]);
            ctObj.wrapText = XmlHelper.ReadBool(node.Attributes["wrapText"]);
            ctObj.indent = XmlHelper.ReadLong(node.Attributes["indent"]);
            ctObj.relativeIndent = XmlHelper.ReadInt(node.Attributes["relativeIndent"]);
            ctObj.justifyLastLine = XmlHelper.ReadBool(node.Attributes["justifyLastLine"]);
            ctObj.shrinkToFit = XmlHelper.ReadBool(node.Attributes["shrinkToFit"]);
            ctObj.readingOrder = XmlHelper.ReadLong(node.Attributes["readingOrder"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            if(this.horizontal != ST_HorizontalAlignment.general)
                XmlHelper.WriteAttribute(sw, "horizontal", this.horizontal.ToString());
            if (this.vertical != ST_VerticalAlignment.bottom)
                XmlHelper.WriteAttribute(sw, "vertical", this.vertical.ToString());
            XmlHelper.WriteAttribute(sw, "textRotation", this.textRotation);
            if(this.wrapText)
                XmlHelper.WriteAttribute(sw, "wrapText", this.wrapText);
            XmlHelper.WriteAttribute(sw, "indent", this.indent);
            XmlHelper.WriteAttribute(sw, "relativeIndent", this.relativeIndent);
            if (justifyLastLine)
                XmlHelper.WriteAttribute(sw, "justifyLastLine", this.justifyLastLine);
            if(shrinkToFit)
                XmlHelper.WriteAttribute(sw, "shrinkToFit", this.shrinkToFit);
            XmlHelper.WriteAttribute(sw, "readingOrder", this.readingOrder);
            sw.Write("/>");
        }

        public bool IsSetHorizontal()
        {
            return this.horizontalFieldSpecified;
        }
        public bool IsSetVertical()
        {
            return this.verticalFieldSpecified;
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
                this.horizontalFieldSpecified = true;
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
        [DefaultValue(ST_VerticalAlignment.bottom)]
        public ST_VerticalAlignment vertical
        {
            get
            {
                return this.verticalField;
            }
            set
            {
                this.verticalField = value;
                this.verticalFieldSpecified = true;
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
                this.textRotationFieldSpecified = true;
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
                this.wrapTextSpecified = true;
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
                this.indentSpecified = true;
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
                this.relativeIndentSpecified = true;
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
                this.justifyLastLineSpecified = true;
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
                this.shrinkToFitSpecified = true;
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
                this.readingOrderSpecified = true;
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

        internal CT_CellAlignment Copy()
        {
            CT_CellAlignment align = new CT_CellAlignment();
            align.horizontal = this.horizontal;
            align.vertical = this.vertical;
            align.wrapText = this.wrapText;
            align.shrinkToFit = this.shrinkToFit;
            align.textRotation = this.textRotation;
            align.justifyLastLine = this.justifyLastLine;
            align.readingOrder = this.readingOrder;
            align.relativeIndent = this.relativeIndent;
            align.indent = this.indent;
            return align;
        }
    }
}
