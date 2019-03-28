using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_PatternFill
    {
        private CT_Color fgColorField = null;

        private CT_Color bgColorField = null;

        private ST_PatternType patternTypeField = ST_PatternType.none;

        public bool IsSetPatternType()
        {
            return this.patternTypeField != ST_PatternType.none;
        }
        public CT_Color AddNewFgColor()
        {
            this.fgColorField = new CT_Color();
            return fgColorField;
        }

        public CT_Color AddNewBgColor()
        {
            this.bgColorField = new CT_Color();
            return bgColorField;
        }
        public void UnsetPatternType()
        {
            this.patternTypeField = ST_PatternType.none;
        }
        public void UnsetFgColor()
        {
            this.fgColorField = null;
        }
        public void UnsetBgColor()
        {
            this.bgColorField = null;
        }
        public static CT_PatternFill Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PatternFill ctObj = new CT_PatternFill();
            if (node.Attributes["patternType"] != null)
                ctObj.patternType = (ST_PatternType)Enum.Parse(typeof(ST_PatternType), node.Attributes["patternType"].Value);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "fgColor")
                    ctObj.fgColor = CT_Color.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bgColor")
                    ctObj.bgColor = CT_Color.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "patternType", this.patternType.ToString());
            if (this.fgColor == null && this.bgColor == null)
            {
                sw.Write("/>");
            }
            else
            {
                sw.Write(">");
                if (this.fgColor != null)
                    this.fgColor.Write(sw, "fgColor");
                if (this.bgColor != null)
                    this.bgColor.Write(sw, "bgColor");
                sw.Write(string.Format("</{0}>", nodeName));
            }
        }

        [XmlElement]
        public CT_Color fgColor
        {
            get
            {
                return this.fgColorField;
            }
            set
            {
                this.fgColorField = value;
            }
        }

        public bool IsSetBgColor()
        {
            return bgColorField != null;
        }

        public bool IsSetFgColor()
        {
            return fgColorField != null;
        }

        [XmlElement]
        public CT_Color bgColor
        {
            get
            {
                return this.bgColorField;
            }
            set
            {
                this.bgColorField = value;
            }
        }

        [XmlAttribute]
        public ST_PatternType patternType
        {
            get
            {
                return this.patternTypeField;
            }
            set
            {
                this.patternTypeField = value;
            }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_PatternType
    {


        none,


        solid,


        mediumGray,


        darkGray,


        lightGray,


        darkHorizontal,


        darkVertical,


        darkDown,


        darkUp,


        darkGrid,


        darkTrellis,


        lightHorizontal,


        lightVertical,


        lightDown,


        lightUp,


        lightGrid,


        lightTrellis,


        gray125,


        gray0625,
    }

}
