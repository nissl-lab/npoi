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
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_BorderPr
    {

        private CT_Color colorField;

        private ST_BorderStyle styleField;
        public static CT_BorderPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_BorderPr ctObj = new CT_BorderPr();
            if (node.Attributes["style"] != null)
                ctObj.style = (ST_BorderStyle)Enum.Parse(typeof(ST_BorderStyle), node.Attributes["style"].Value);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "color")
                    ctObj.color = CT_Color.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            if(this.style!= ST_BorderStyle.none)
                XmlHelper.WriteAttribute(sw, "style", this.style.ToString());
            if (this.color != null)
            {
                sw.Write(">");
                this.color.Write(sw, "color");
                sw.Write(string.Format("</{0}>", nodeName));
            }
            else
            {
                sw.Write("/>");
            }
        }


        public CT_BorderPr()
        {
            //this.colorField = new CT_Color();
            this.styleField = ST_BorderStyle.none;
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

        public CT_BorderPr Copy()
        {
            var res = new CT_BorderPr();
            res.colorField = this.colorField == null ? null : this.colorField.Copy();
            res.style = this.style;
            return res;
        }
    }
}
