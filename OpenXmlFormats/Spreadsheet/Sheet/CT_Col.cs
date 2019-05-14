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
        private double widthField;  // optional, but width has no default value
        private bool collapsedField = true;

        [XmlAttribute]
        public uint min { get; set; }

        [XmlAttribute]
        public uint max { get; set; }

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
        public bool widthSpecified { get; set; }


        public bool IsSetBestFit()
        {
            return this.bestFit != false;
        }
        public bool IsSetCustomWidth()
        {
            return this.customWidth != false;
        }
        public bool IsSetHidden()
        {
            return this.hidden != false;
        }
        public bool IsSetStyle()
        {
            return this.style!=null;
        }
        public bool IsSetWidth()
        {
            return this.widthField > 0;
        }
        public bool IsSetCollapsed()
        {
            return this.collapsedSpecified;
        }
        public bool IsSetPhonetic()
        {
            return this.phonetic != false;
        }
        public bool IsSetOutlineLevel()
        {
            return this.outlineLevel != 0;
        }
        public void UnsetHidden()
        {
            this.hidden = false;
        }
        public void UnsetCollapsed()
        {
            this.collapsedField = true;
            this.collapsedSpecified = false;
        }



        [XmlAttribute]
        public uint? style { get; set; }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool hidden { get; set; }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool bestFit { get; set; }
        [DefaultValue(false)]
        [XmlAttribute]
        public bool customWidth { get; set; }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool phonetic { get; set; }


        [DefaultValue(typeof(byte), "0")]
        [XmlAttribute]
        public byte outlineLevel { get; set; }

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
                this.collapsedSpecified = true;
            }
        }
        [XmlIgnore]
        public bool collapsedSpecified { get; set; } = true;

        public CT_Col()
        {
            
        }

        public static CT_Col Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Col ctObj = new CT_Col();
            ctObj.min = XmlHelper.ReadUInt(node.Attributes[nameof(min)]);
            ctObj.max = XmlHelper.ReadUInt(node.Attributes[nameof(max)]);
            ctObj.width = XmlHelper.ReadDouble(node.Attributes[nameof(width)]);
            if (node.Attributes[nameof(style)] != null)
                ctObj.style = XmlHelper.ReadUInt(node.Attributes[nameof(style)]);
            else
                ctObj.style = null;
            ctObj.hidden = XmlHelper.ReadBool(node.Attributes[nameof(hidden)]);
            ctObj.bestFit = XmlHelper.ReadBool(node.Attributes[nameof(bestFit)]);
            ctObj.outlineLevel = XmlHelper.ReadByte(node.Attributes[nameof(outlineLevel)]);
            ctObj.customWidth = XmlHelper.ReadBool(node.Attributes[nameof(customWidth)]);
            ctObj.phonetic = XmlHelper.ReadBool(node.Attributes[nameof(phonetic)]);
            ctObj.collapsed = XmlHelper.ReadBool(node.Attributes[nameof(collapsed)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(min), this.min);
            XmlHelper.WriteAttribute(sw, nameof(max), this.max);
            XmlHelper.WriteAttribute(sw, nameof(width), this.width);
            if(this.style!=null)
                XmlHelper.WriteAttribute(sw, nameof(style), (uint)this.style,true);
            XmlHelper.WriteAttribute(sw, nameof(hidden), this.hidden,false);
            XmlHelper.WriteAttribute(sw, nameof(bestFit), this.bestFit,false);
            XmlHelper.WriteAttribute(sw, nameof(customWidth), this.customWidth,false);
            XmlHelper.WriteAttribute(sw, nameof(phonetic), this.phonetic,false);
            XmlHelper.WriteAttribute(sw, nameof(outlineLevel), this.outlineLevel);
            XmlHelper.WriteAttribute(sw, nameof(collapsed), this.collapsed,false);
            sw.Write("/>");
        }





        public CT_Col Copy()
        {
            CT_Col col = new CT_Col();
            col.bestFit = this.bestFit;
            col.collapsedField = this.collapsedField;
            col.collapsedSpecified = this.collapsedSpecified;
            col.customWidth = this.customWidth;
            col.hidden = this.hidden;
            col.max = this.max;
            col.min = this.min;
            col.outlineLevel = this.outlineLevel;
            col.phonetic = this.phonetic;
            col.style = this.style;
            col.widthField = this.widthField;
            col.widthSpecified = this.widthSpecified;
            
            return col;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;
            if (!(obj is CT_Col))
                return false;
            CT_Col col = obj as CT_Col;
            return (col.min == this.min) && (col.max == this.max);
        }

        public override int GetHashCode()
        {
            return this.ToString().GetHashCode();
        }

        public override string ToString()
        {
            return $"min:{this.min}, max:{this.max}, width:{this.width}";
        }
    }
}
