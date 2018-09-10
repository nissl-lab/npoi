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
    public class CT_Colors
    {
        private List<CT_RgbColor> indexedColorsField;

        private List<CT_Color> mruColorsField;

        public CT_Colors()
        {
            //this.mruColorsField = new List<CT_Color>();
            //this.indexedColorsField = new List<CT_RgbColor>();
        }
        public static CT_Colors Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Colors ctObj = new CT_Colors();
            
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "indexedColors")
                {
                    ctObj.indexedColors = new List<CT_RgbColor>();
                    foreach (XmlNode c2Node in childNode.ChildNodes)
                    {
                        ctObj.indexedColors.Add(CT_RgbColor.Parse(c2Node, namespaceManager));
                    }
                }
                else if (childNode.LocalName == "mruColors")
                {
                    ctObj.mruColors = new List<CT_Color>();
                    foreach (XmlNode c2Node in childNode.ChildNodes)
                    {
                        ctObj.mruColors.Add(CT_Color.Parse(c2Node, namespaceManager));
                    }
                }
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}>", nodeName));
            if (this.indexedColors != null)
            {
                sw.Write("<indexedColors>");
                foreach (CT_RgbColor x in this.indexedColors)
                {
                    x.Write(sw, "rgbColor");
                }
                sw.Write("</indexedColors>");
            }
            if (this.mruColors != null)
            {
                sw.Write("<mruColors>");
                foreach (CT_Color x in this.mruColors)
                {
                    x.Write(sw, "color");
                }
                sw.Write("</mruColors>");
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlArray(Order = 0)]
        [XmlArrayItem("rgbColor", IsNullable = false)]
        public List<CT_RgbColor> indexedColors
        {
            get
            {
                return this.indexedColorsField;
            }
            set
            {
                this.indexedColorsField = value;
            }
        }

        [XmlArray(Order = 1)]
        [XmlArrayItem("color", IsNullable = false)]
        public List<CT_Color> mruColors
        {
            get
            {
                return this.mruColorsField;
            }
            set
            {
                this.mruColorsField = value;
            }
        }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_RgbColor
    {
        public static CT_RgbColor Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RgbColor ctObj = new CT_RgbColor();
            ctObj.rgb = XmlHelper.ReadBytes(node.Attributes["rgb"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "rgb", this.rgb);
            sw.Write("/>");
        }

        private byte[] rgbField = null; // ARGB

        [XmlAttribute(DataType = "hexBinary")]
        // Type ST_UnsignedIntHex is base on xsd:hexBinary, Length 4 (octets!?)
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
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("color", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Color
    {
        // all attributes are optional
        private bool autoField;

        private uint indexedField;

        private byte[] rgbField; // type ST_UnsignedIntHex is xsd:hexBinary restricted to length 4 (octets!? - see http://www.grokdoc.net/index.php/EOOXML_Objections_Clearinghouse)

        private uint themeField; // TODO change all the uses theme to use uint instead of signed integer variants

        private double tintField;


        #region auto
        [XmlAttribute]
        public bool auto
        {
            get
            {
                return (bool)this.autoField;
            }
            set
            {
                this.autoField = value;
                this.autoSpecified = true;
            }
        }
        bool autoSpecifiedField = false;
        [XmlIgnore]
        // do not remove this field or change the name, because it is automatically used by the XmlSerializer to decide if the auto attribute should be printed or not.
        public bool autoSpecified
        {
            get { return autoSpecifiedField; }
            set { autoSpecifiedField = value; }
        }
        public bool IsSetAuto()
        {
            return autoSpecifiedField;
        }

        #endregion auto

        #region indexed
        [XmlAttribute]
        public uint indexed
        {
            get
            {
                return (uint)this.indexedField;
            }
            set
            {
                this.indexedField = value;
                this.indexedSpecifiedField = true;
            }
        }
        bool indexedSpecifiedField = false;
        [XmlIgnore]
        public bool indexedSpecified
        {
            get { return indexedSpecifiedField; }
            set { indexedSpecifiedField = value; }
        }
        public bool IsSetIndexed()
        {
            return indexedSpecified;
        }
        #endregion indexed

        #region rgb
        [XmlAttribute(DataType = "hexBinary")]
        // Type ST_UnsignedIntHex is base on xsd:hexBinary
        public byte[] rgb
        {
            get
            {
                return this.rgbField;
            }
            set
            {
                this.rgbField = value;
                this.rgbSpecified = true;
            }
        }
        bool rgbSpecifiedField = false;
        [XmlIgnore]
        public bool rgbSpecified
        {
            get { return rgbSpecifiedField; }
            set { rgbSpecifiedField = value; }
        }
        public void SetRgb(byte R, byte G, byte B)
        {
            this.rgbField = new byte[4];
            this.rgbField[0] = 0;
            this.rgbField[1] = R;
            this.rgbField[2] = G;
            this.rgbField[3] = B;
            rgbSpecified = true;
        }
        public bool IsSetRgb()
        {
            return rgbSpecified;
        }
        public void SetRgb(byte[] rgb)
        {
            rgbField = new byte[rgb.Length];
            Array.Copy(rgb, this.rgbField, rgb.Length);
            this.rgbSpecified = true;
        }
        public byte[] GetRgb()
        {
            if (rgbField == null) return null;
            byte[] retVal = new byte[rgbField.Length];
            Array.Copy(rgbField, retVal, rgbField.Length);
            return retVal;
        }
        #endregion rgb

        #region theme
        [XmlAttribute]
        public uint theme
        {
            get
            {
                return (uint)this.themeField;
            }
            set
            {
                this.themeField = value;
                this.themeSpecifiedField = true;
            }
        }
        bool themeSpecifiedField;
        [XmlIgnore]
        public bool themeSpecified
        {
            get { return themeSpecifiedField; }
            set { themeSpecifiedField = value; }

        }
        public bool IsSetTheme()
        {
            return themeSpecified;
        }
        #endregion theme

        #region tint
        [DefaultValue(0.0D)]
        [XmlAttribute]
        public double tint
        {
            get
            {
                return this.tintField;
            }
            set
            {
                this.tintField = value;
                this.tintSpecified = true;
            }
        }
        bool tintSpecifiedField = false;
        [XmlIgnore]
        public bool tintSpecified
        {
            get { return tintSpecifiedField; }
            set { tintSpecifiedField = value; }

        }
        public bool IsSetTint()
        {
            return tintSpecified;
        }
        #endregion tint

        //internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Color));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });
        //public override string ToString()
        //{
        //    using (StringWriter stringWriter = new StringWriter())
        //    {
        //        serializer.Serialize(stringWriter, this, namespaces);
        //        return stringWriter.ToString();
        //    }
        //}

        public static CT_Color Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Color ctObj = new CT_Color();
            ctObj.auto = XmlHelper.ReadBool(node.Attributes["auto"]);
            ctObj.autoSpecified = node.Attributes["auto"] != null;
            ctObj.indexed = XmlHelper.ReadUInt(node.Attributes["indexed"]);
            ctObj.indexedSpecified = node.Attributes["indexed"] != null;
            ctObj.rgb = XmlHelper.ReadBytes(node.Attributes["rgb"]);
            ctObj.rgbSpecified = node.Attributes["rgb"] != null;
            ctObj.theme = XmlHelper.ReadUInt(node.Attributes["theme"]);
            ctObj.themeSpecified = node.Attributes["theme"] != null;
            ctObj.tint = XmlHelper.ReadDouble(node.Attributes["tint"]);
            ctObj.tintSpecified = node.Attributes["tint"] != null;
            return ctObj;
        }




        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "auto", this.auto,false);
            if (indexedSpecified)
                XmlHelper.WriteAttribute(sw, "indexed", this.indexed, true);
            if(rgbSpecified)
                XmlHelper.WriteAttribute(sw, "rgb", this.rgb);
            if (themeSpecified)
                XmlHelper.WriteAttribute(sw, "theme", this.theme, true);
            if(tintSpecified)
                XmlHelper.WriteAttribute(sw, "tint", this.tint);
            sw.Write("/>");
        }

        public CT_Color Copy()
        {
            var res = new CT_Color();
            res.autoField = this.autoField;

            res.indexedField = this.indexedField;

            res.rgbField = this.rgbField == null ? null : (byte[])this.rgbField.Clone(); // type ST_UnsignedIntHex is xsd:hexBinary restricted to length 4 (octets!? - see http://www.grokdoc.net/index.php/EOOXML_Objections_Clearinghouse)

            res.themeField = this.themeField; // TODO change all the uses theme to use uint instead of signed integer variants

            res.tintField = this.tintField;

            return res;
        }
    }

}
