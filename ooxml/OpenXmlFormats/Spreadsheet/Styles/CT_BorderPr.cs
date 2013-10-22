using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_BorderPr
    {

        private CT_Color colorField;

        private ST_BorderStyle styleField;

        public CT_BorderPr()
        {
            //this.colorField = new CT_Color();
            //this.styleField = ST_BorderStyle.none;
        }
        public void SetColor(CT_Color color)
        {
            this.colorField = color;
        }
        public bool IsSetColor()
        {
            return colorField != null;
        }
        public void UnsetColor()
        {
            colorField = null;
        }
        public bool IsSetStyle()
        {
            return styleField != ST_BorderStyle.none;
        }

        [XmlElement]
        public CT_Color color
        {
            get
            {
                return this.colorField;
            }
            set
            {
                this.colorField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(ST_BorderStyle.none)]
        public ST_BorderStyle style
        {
            get
            {
                return this.styleField;
            }
            set
            {
                this.styleField = value;
            }
        }
    }
}
