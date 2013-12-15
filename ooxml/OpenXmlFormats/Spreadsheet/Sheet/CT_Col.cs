using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Text;
using System.Xml.Serialization;
using System.Xml;
using NPOI.OpenXml4Net.Util;
using System.IO;

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
            return this.styleField!=null;
        }
        public bool IsSetWidth()
        {
            return this.widthField>0;
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

        public CT_Col()
        {
            
        }

        public static CT_Col Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Col ctObj = new CT_Col();
            ctObj.min = XmlHelper.ReadUInt(node.Attributes["min"]);
            ctObj.max = XmlHelper.ReadUInt(node.Attributes["max"]);
            ctObj.width = XmlHelper.ReadDouble(node.Attributes["width"]);
            ctObj.style = XmlHelper.ReadUInt(node.Attributes["style"]);
            ctObj.hidden = XmlHelper.ReadBool(node.Attributes["hidden"]);
            ctObj.bestFit = XmlHelper.ReadBool(node.Attributes["bestFit"]);
            ctObj.customWidth = XmlHelper.ReadBool(node.Attributes["customWidth"]);
            ctObj.phonetic = XmlHelper.ReadBool(node.Attributes["phonetic"]);
            ctObj.collapsed = XmlHelper.ReadBool(node.Attributes["collapsed"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "min", this.min);
            XmlHelper.WriteAttribute(sw, "max", this.max);
            XmlHelper.WriteAttribute(sw, "width", this.width);
            XmlHelper.WriteAttribute(sw, "style", this.style,true);
            XmlHelper.WriteAttribute(sw, "hidden", this.hidden,false);
            XmlHelper.WriteAttribute(sw, "bestFit", this.bestFit,false);
            XmlHelper.WriteAttribute(sw, "customWidth", this.customWidth,false);
            XmlHelper.WriteAttribute(sw, "phonetic", this.phonetic,false);
            XmlHelper.WriteAttribute(sw, "outlineLevel", this.outlineLevel);
            XmlHelper.WriteAttribute(sw, "collapsed", this.collapsed,false);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }




    }
}
