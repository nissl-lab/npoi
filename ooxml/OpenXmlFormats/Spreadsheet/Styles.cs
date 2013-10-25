
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Spreadsheet
{














    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(ElementName = "styleSheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class CT_Stylesheet
    {

        private CT_NumFmts numFmtsField;

        private CT_Fonts fontsField;

        private CT_Fills fillsField;

        private CT_Borders bordersField;

        private CT_CellStyleXfs cellStyleXfsField;

        private CT_CellXfs cellXfsField;

        private CT_CellStyles cellStylesField;

        private CT_Dxfs dxfsField;

        private CT_TableStyles tableStylesField;

        private CT_Colors colorsField;

        private CT_ExtensionList extLstField;

        public CT_Stylesheet()
        {
            //this.extLstField = new CT_ExtensionList();
            //this.colorsField = new CT_Colors();
            //this.tableStylesField = new CT_TableStyles();
            //this.dxfsField = new CT_Dxfs();
            //this.cellStylesField = new CT_CellStyles();
            //this.bordersField = new CT_Borders();
            //this.fillsField = new CT_Fills();
            //this.fontsField = new CT_Fonts();
            //this.numFmtsField = new CT_NumFmts();
        }
        public CT_Borders AddNewBorders()
        {
            this.bordersField = new CT_Borders();
            return this.bordersField;
        }
        public CT_CellStyleXfs AddNewCellStyleXfs()
        {
            this.cellStyleXfsField = new CT_CellStyleXfs();
            return this.cellStyleXfsField;
        }
        public CT_CellXfs AddNewCellXfs()
        {
            this.cellXfsField = new CT_CellXfs();
            return this.cellXfsField;
        }
        [XmlElement]
        public CT_NumFmts numFmts
        {
            get
            {
                return this.numFmtsField;
            }
            set
            {
                this.numFmtsField = value;
            }
        }
        [XmlElement]
        public CT_Fonts fonts
        {
            get
            {
                return this.fontsField;
            }
            set
            {
                this.fontsField = value;
            }
        }
        [XmlElement]
        public CT_Fills fills
        {
            get
            {
                return this.fillsField;
            }
            set
            {
                this.fillsField = value;
            }
        }
        [XmlElement]
        public CT_Borders borders
        {
            get
            {
                return this.bordersField;
            }
            set
            {
                this.bordersField = value;
            }
        }
        [XmlElement]
        public CT_CellStyleXfs cellStyleXfs
        {
            get
            {
                return this.cellStyleXfsField;
            }
            set
            {
                this.cellStyleXfsField = value;
            }
        }
        [XmlElement]
        public CT_CellXfs cellXfs
        {
            get
            {
                return this.cellXfsField;
            }
            set
            {
                this.cellXfsField = value;
            }
        }
        [XmlElement]
        public CT_CellStyles cellStyles
        {
            get
            {
                return this.cellStylesField;
            }
            set
            {
                this.cellStylesField = value;
            }
        }
        [XmlElement]
        public CT_Dxfs dxfs
        {
            get
            {
                return this.dxfsField;
            }
            set
            {
                this.dxfsField = value;
            }
        }
        [XmlElement]
        public CT_TableStyles tableStyles
        {
            get
            {
                return this.tableStylesField;
            }
            set
            {
                this.tableStylesField = value;
            }
        }
        [XmlElement]
        public CT_Colors colors
        {
            get
            {
                return this.colorsField;
            }
            set
            {
                this.colorsField = value;
            }
        }
        [XmlElement]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }
    }


    public class CT_Dxf
    {

        private CT_Font fontField;

        private CT_NumFmt numFmtField;

        private CT_Fill fillField;

        private CT_CellAlignment alignmentField;

        private CT_Border borderField;

        private CT_CellProtection protectionField;

        private CT_ExtensionList extLstField;

        public CT_Dxf()
        {
            //this.extLstField = new CT_ExtensionList();
            //this.protectionField = new CT_CellProtection();
            //this.borderField = new CT_Border();
            //this.alignmentField = new CT_CellAlignment();
            //this.fillField = new CT_Fill();
            //this.numFmtField = new CT_NumFmt();
            //this.fontField = new CT_Font();
        }

        public bool IsSetBorder()
        {
            return borderField != null;
        }
        public CT_Font AddNewFont()
        {
            CT_Font font = new CT_Font();
            this.fontField = font;
            return font;
        }

        public CT_Fill AddNewFill()
        {
            CT_Fill fill = new CT_Fill();
            this.fillField = fill;
            return fill;
        }

        public CT_Border AddNewBorder()
        {
            CT_Border border = new CT_Border();
            this.borderField = border;
            return border;
        }
        public bool IsSetFont()
        {
            return fontField != null;
        }
        public bool IsSetFill()
        {
            return fillField != null;
        }
        [XmlElement]
        public CT_Font font
        {
            get
            {
                return this.fontField;
            }
            set
            {
                this.fontField = value;
            }
        }
        [XmlElement]
        public CT_NumFmt numFmt
        {
            get
            {
                return this.numFmtField;
            }
            set
            {
                this.numFmtField = value;
            }
        }
        [XmlElement]
        public CT_Fill fill
        {
            get
            {
                return this.fillField;
            }
            set
            {
                this.fillField = value;
            }
        }
        [XmlElement]
        public CT_CellAlignment alignment
        {
            get
            {
                return this.alignmentField;
            }
            set
            {
                this.alignmentField = value;
            }
        }
        [XmlElement]
        public CT_Border border
        {
            get
            {
                return this.borderField;
            }
            set
            {
                this.borderField = value;
            }
        }
        [XmlElement]
        public CT_CellProtection protection
        {
            get
            {
                return this.protectionField;
            }
            set
            {
                this.protectionField = value;
            }
        }
        [XmlElement]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_RgbColor
    {

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
    public class CT_Colors
    {
        private List<CT_RgbColor> indexedColorsField;

        private List<CT_Color> mruColorsField;

        public CT_Colors()
        {
            this.mruColorsField = new List<CT_Color>();
            this.indexedColorsField = new List<CT_RgbColor>();
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
    public enum ST_GradientType
    {
        NONE,
    
        linear,

    
        path,
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_PatternFill
    {
        private CT_Color fgColorField = null;

        private CT_Color bgColorField = null;

        private ST_PatternType? patternTypeField = null;

        public bool IsSetPatternType()
        {
            return this.patternTypeField != null;
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
        public void unsetPatternType()
        {
            this.patternTypeField = null;
        }
        public void unsetFgColor()
        {
            this.fgColorField = null;
        }
        public void unsetBgColor()
        {
            this.bgColorField = null;
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
                return null != this.patternTypeField ? (ST_PatternType)this.patternTypeField : ST_PatternType.none;
            }
            set
            {
                this.patternTypeField = value;
            }
        }

        [XmlIgnore]
        public bool patternTypeSpecified
        {
            get
            {
                return null != this.patternTypeField;
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

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("color",Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
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
                this.indexedSpecifiedField = indexed != 0;
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
            return rgbField;
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

        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Color));
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });
        public override string ToString()
        {
            using (StringWriter stringWriter = new StringWriter())
            {
                serializer.Serialize(stringWriter, this, namespaces);
                return stringWriter.ToString();
            }
        }

        public static CT_Color Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Color ctObj = new CT_Color();
            ctObj.auto = XmlHelper.ReadBool(node.Attributes["auto"]);
            ctObj.indexed = XmlHelper.ReadUInt(node.Attributes["indexed"]);
            ctObj.rgb = XmlHelper.ReadBytes(node.Attributes["rgb"]);
            ctObj.theme = XmlHelper.ReadUInt(node.Attributes["theme"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "auto", this.auto);
            XmlHelper.WriteAttribute(sw, "indexed", this.indexed);
            XmlHelper.WriteAttribute(sw, "rgb", this.rgb);
            XmlHelper.WriteAttribute(sw, "theme", this.theme);
            XmlHelper.WriteAttribute(sw, "tint", this.tint);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }


    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_FontScheme
    {
        private ST_FontScheme valField;

    
        [XmlAttribute]
        public ST_FontScheme val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }

        public static CT_FontScheme Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FontScheme ctObj = new CT_FontScheme();
            if (node.Attributes["val"] != null)
                ctObj.val = (ST_FontScheme)Enum.Parse(typeof(ST_FontScheme), node.Attributes["val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }


    }



    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_FontName
    {

        private string valField;

        public static CT_FontName Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FontName ctObj = new CT_FontName();
            ctObj.val = XmlHelper.ReadString(node.Attributes["val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }


        [XmlAttribute]
        public string val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_FontSize
    {
        private double valField;
        public static CT_FontSize Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FontSize ctObj = new CT_FontSize();
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }


        [XmlAttribute]
        public double val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_FontScheme
    {
    
        none = 1,

    
        major = 2,

    
        minor = 3,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_UnderlineProperty
    {

        private ST_UnderlineValues? valField = null;

        public CT_UnderlineProperty()
        {
            this.valField = ST_UnderlineValues.single;
        }

        public static CT_UnderlineProperty Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_UnderlineProperty ctObj = new CT_UnderlineProperty();
            if (node.Attributes["val"] != null)
                ctObj.val = (ST_UnderlineValues)Enum.Parse(typeof(ST_UnderlineValues), node.Attributes["val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }




        [DefaultValue(ST_UnderlineValues.single)]
        [XmlAttribute]
        public ST_UnderlineValues val
        {
            get
            {
                return  (null == valField) ? ST_UnderlineValues.single : (ST_UnderlineValues)this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
        [XmlIgnore]
        public bool valbSpecified
        {
            get { return (null != valField); }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_UnderlineValues
    {
    
        none,

    
        single,

    
        [XmlEnum("double")]
        @double,

    
        singleAccounting,

    
        doubleAccounting,

    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_VerticalAlignFontProperty
    {
        private ST_VerticalAlignRun valField;

        [XmlAttribute]
        public ST_VerticalAlignRun val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
        public static CT_VerticalAlignFontProperty Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_VerticalAlignFontProperty ctObj = new CT_VerticalAlignFontProperty();
            if (node.Attributes["val"] != null)
                ctObj.val = (ST_VerticalAlignRun)Enum.Parse(typeof(ST_VerticalAlignRun), node.Attributes["val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }


    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_VerticalAlignRun
    {
    
        baseline,

    
        superscript,

    
        subscript,
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_BooleanProperty
    {

        private bool? valField = null;

        public CT_BooleanProperty()
        {
            this.valField = true;
        }
        public static CT_BooleanProperty Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_BooleanProperty ctObj = new CT_BooleanProperty();
            ctObj.val = XmlHelper.ReadBool(node.Attributes["val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }


        [DefaultValue(true)]
        [XmlAttribute]
        public bool val
        {
            get
            {
                return null == this.valField ? true : (bool)this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
        [XmlIgnore]
        public bool valSpecified
        {
            get { return (null != valField); }
        }

    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_IntProperty
    {

        private int valField;

        [XmlAttribute]
        public int val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }

        public static CT_IntProperty Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_IntProperty ctObj = new CT_IntProperty();
            ctObj.val = XmlHelper.ReadInt(node.Attributes["val"]);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }
    }


    public enum ST_TableStyleType
    {

    
        wholeTable,

    
        headerRow,

    
        totalRow,

    
        firstColumn,

    
        lastColumn,

    
        firstRowStripe,

    
        secondRowStripe,

    
        firstColumnStripe,

    
        secondColumnStripe,

    
        firstHeaderCell,

    
        lastHeaderCell,

    
        firstTotalCell,

    
        lastTotalCell,

    
        firstSubtotalColumn,

    
        secondSubtotalColumn,

    
        thirdSubtotalColumn,

    
        firstSubtotalRow,

    
        secondSubtotalRow,

    
        thirdSubtotalRow,

    
        blankRow,

    
        firstColumnSubheading,

    
        secondColumnSubheading,

    
        thirdColumnSubheading,

    
        firstRowSubheading,

    
        secondRowSubheading,

    
        thirdRowSubheading,

    
        pageFieldLabels,

    
        pageFieldValues,
    }

    public class CT_TableStyle
    {

        private List<CT_TableStyleElement> tableStyleElementField;

        private string nameField;

        private bool pivotField;

        private bool tableField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_TableStyle()
        {
            this.tableStyleElementField = new List<CT_TableStyleElement>();
            this.pivotField = true;
            this.tableField = true;
        }
        [XmlElement]
        public List<CT_TableStyleElement> tableStyleElement
        {
            get
            {
                return this.tableStyleElementField;
            }
            set
            {
                this.tableStyleElementField = value;
            }
        }
        [XmlAttribute]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }
        [XmlAttribute]
        [DefaultValue(true)]
        public bool pivot
        {
            get
            {
                return this.pivotField;
            }
            set
            {
                this.pivotField = value;
            }
        }
        [XmlAttribute]
        [DefaultValue(true)]
        public bool table
        {
            get
            {
                return this.tableField;
            }
            set
            {
                this.tableField = value;
            }
        }
        [XmlAttribute]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [XmlIgnore]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }

    public class CT_TableStyles
    {

        private List<CT_TableStyle> tableStyleField;

        private uint countField;

        private bool countFieldSpecified;

        private string defaultTableStyleField;

        private string defaultPivotStyleField;

        public CT_TableStyles()
        {
            this.tableStyleField = new List<CT_TableStyle>();
        }
        [XmlElement]
        public List<CT_TableStyle> tableStyle
        {
            get
            {
                return this.tableStyleField;
            }
            set
            {
                this.tableStyleField = value;
            }
        }
        [XmlAttribute]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [XmlIgnore]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }

        public string defaultTableStyle
        {
            get
            {
                return this.defaultTableStyleField;
            }
            set
            {
                this.defaultTableStyleField = value;
            }
        }

        public string defaultPivotStyle
        {
            get
            {
                return this.defaultPivotStyleField;
            }
            set
            {
                this.defaultPivotStyleField = value;
            }
        }
    }
    public class CT_TableStyleElement
    {

        private ST_TableStyleType typeField;

        private uint sizeField;

        private uint dxfIdField;

        private bool dxfIdFieldSpecified;

        public CT_TableStyleElement()
        {
            this.sizeField = ((uint)(1));
        }
        [XmlAttribute]
        public ST_TableStyleType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [DefaultValue(typeof(uint), "1")]
        public uint size
        {
            get
            {
                return this.sizeField;
            }
            set
            {
                this.sizeField = value;
            }
        }
        [XmlAttribute]
        public uint dxfId
        {
            get
            {
                return this.dxfIdField;
            }
            set
            {
                this.dxfIdField = value;
            }
        }

        [XmlIgnore]
        public bool dxfIdSpecified
        {
            get
            {
                return this.dxfIdFieldSpecified;
            }
            set
            {
                this.dxfIdFieldSpecified = value;
            }
        }
    }
    public class CT_Dxfs
    {

        private List<CT_Dxf> dxfField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_Dxfs()
        {
            this.dxfField = new List<CT_Dxf>();
        }
        [XmlElement]
        public List<CT_Dxf> dxf
        {
            get
            {
                return this.dxfField;
            }
            set
            {
                this.dxfField = value;
            }
        }
        [XmlAttribute]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [XmlIgnore]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }
}
