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

        private uint? styleField;// optional, as are the following attributes

        private bool hiddenField;

        private bool bestFitField;

        private bool customWidthField;

        private bool phoneticField;

        private byte outlineLevelField;

        private bool collapsedField = false;
        private bool collapsedSpecifiedField = false;

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
            return this.styleField != null && this.styleField != 0;
        }
        public bool IsSetWidth()
        {
            return this.widthField > 0;
        }

        public bool IsSetNumber()
        {
            return min > 0 && max > 0;
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

        public void SetNumber(uint number)
        {
            min = number;
            max = number;
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

        public void UnsetWidth()
        {
            this.widthField = -1;
        }

        public void UnsetCustomWidth()
        {
            this.customWidthField = false;
        }

        public void UnsetStyle()
        {
            this.styleField = 0;
        }

        [XmlAttribute]
        public uint? style
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
            ctObj.widthField = XmlHelper.ReadDouble(node.Attributes["width"]);
            if (node.Attributes["style"] != null)
                ctObj.style = XmlHelper.ReadUInt(node.Attributes["style"]);
            else
                ctObj.style = null;
            ctObj.hidden = XmlHelper.ReadBool(node.Attributes["hidden"]);
            ctObj.bestFit = XmlHelper.ReadBool(node.Attributes["bestFit"]);
            ctObj.outlineLevel = XmlHelper.ReadByte(node.Attributes["outlineLevel"]);
            ctObj.customWidthField = XmlHelper.ReadBool(node.Attributes["customWidth"]);
            ctObj.phonetic = XmlHelper.ReadBool(node.Attributes["phonetic"]);
            ctObj.collapsedField = XmlHelper.ReadBool(node.Attributes["collapsed"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "min", this.min);
            XmlHelper.WriteAttribute(sw, "max", this.max);
            XmlHelper.WriteAttribute(sw, "width", this.width);
            if(this.style!=null)
                XmlHelper.WriteAttribute(sw, "style", (uint)this.style,true);
            XmlHelper.WriteAttribute(sw, "hidden", this.hidden,false);
            XmlHelper.WriteAttribute(sw, "bestFit", this.bestFit,false);
            XmlHelper.WriteAttribute(sw, "customWidth", this.customWidth,false);
            XmlHelper.WriteAttribute(sw, "phonetic", this.phonetic,false);
            XmlHelper.WriteAttribute(sw, "outlineLevel", this.outlineLevel);
            XmlHelper.WriteAttribute(sw, "collapsed", this.collapsed,false);
            sw.Write("/>");
        }

        public void Set(CT_Col col)
        {
            bestFitField = col.bestFitField;
            collapsedField = col.collapsedField;
            collapsedSpecifiedField = col.collapsedSpecifiedField;
            customWidthField = col.customWidthField;
            hiddenField = col.hiddenField;
            maxField = col.maxField;
            minField = col.minField;
            outlineLevelField = col.outlineLevelField;
            phoneticField = col.phoneticField;
            styleField = col.styleField;
            widthField = col.widthField;
            widthSpecifiedField = col.widthSpecifiedField;
        }

        public CT_Col Copy()
        {
            CT_Col col = new CT_Col();
            col.bestFitField = this.bestFitField;
            col.collapsedField = this.collapsedField;
            col.collapsedSpecifiedField = this.collapsedSpecifiedField;
            col.customWidthField = this.customWidthField;
            col.hiddenField = this.hiddenField;
            col.maxField = this.maxField;
            col.minField = this.minField;
            col.outlineLevelField = this.outlineLevelField;
            col.phoneticField = this.phoneticField;
            col.styleField = this.styleField;
            col.widthField = this.widthField;
            col.widthSpecifiedField = this.widthSpecifiedField;
            
            return col;
        }

        /// <summary>
        /// Checks if <paramref name="col"/> is adjacent or intersecting with
        /// current column and can be combined with it. If that is the case -
        /// will modify current <see cref="CT_Col"/> to have <see cref="min"/>
        /// and <see cref="max"/> of both columns.
        /// </summary>
        /// <param name="col"><see cref="CT_Col"/> to combine with</param>
        /// <exception cref="InvalidOperationException">Thrown if
        /// <paramref name="col"/> cannot be combined with current
        /// <see cref="CT_Col"/></exception>
        public void CombineWith(CT_Col col)
        {
            if(!IsAdjacentAndCanBeCombined(col))
            {
                throw new InvalidOperationException($"Columns [{this}] and " +
                    $"[{col}] could not be combined");
            }

            min = Math.Min(min, col.min);
            max = Math.Max(max, col.max);
        }

        /// <summary>
        /// Checks if <paramref name="col"/> is adjacent or intersecting with
        /// current column and can be combined with it.
        /// </summary>
        /// <param name="col"> <see cref="CT_Col"/> to check</param>
        /// <returns><c>true</c> if <paramref name="col"/> is adjacent or
        /// intersecting and have equal properties with current
        /// <see cref="CT_Col"/>; <c>false</c> otherwise</returns>
        public bool IsAdjacentAndCanBeCombined(CT_Col col)
        {
            return IsAdjacentOrIntersecting(col) && HasEqualProperties(col);
        }

        private bool HasEqualProperties(CT_Col col)
        {
            return bestFitField == col.bestFitField
                && collapsedField == col.collapsedField
                && collapsedSpecifiedField == col.collapsedSpecifiedField
                && customWidthField == col.customWidthField
                && hiddenField == col.hiddenField
                && outlineLevelField == col.outlineLevelField
                && phoneticField == col.phoneticField
                && styleField == col.styleField
                && widthField == col.widthField
                && widthSpecifiedField == col.widthSpecifiedField;
        }

        private bool IsAdjacentOrIntersecting(CT_Col col)
        {
            return IsAdjacent(col) || IsIntersecting(col);
        }

        private bool IsAdjacent(CT_Col col)
        {
            return min == col.max + 1 || max == col.min - 1;
        }

        private bool IsIntersecting(CT_Col col)
        {
            return (min >= col.min && min <= col.max) || (max <= col.max && max >= col.min) ||
                (col.min >= min && col.min <= max) || (col.max <= max && col.max >= min);
        }

        public override bool Equals(object obj)
        {
            if (obj is not CT_Col col)
            {
                return false;
            }

            return col.min == min && col.max == max && HasEqualProperties(col);
        }

        public override int GetHashCode()
        {
            return this.ToString().GetHashCode();
        }

        public override string ToString()
        {
            return string.Format("min:{0}, max:{1}, width:{2}", this.min, this.max, this.width);
        }
    }
}
