using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Fonts
    {

        private List<CT_Font> fontField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_Fonts()
        {
            this.fontField = new List<CT_Font>();
        }

        public void SetFontArray(CT_Font[] array)
        {
            fontField = new List<CT_Font>(array);
        }
        [XmlElement]
        public List<CT_Font> font
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

        [System.Xml.Serialization.XmlIgnore]
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
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRoot(ElementName = "styleSheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
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
            this.extLstField = new CT_ExtensionList();
            this.colorsField = new CT_Colors();
            this.tableStylesField = new CT_TableStyles();
            this.dxfsField = new CT_Dxfs();
            this.cellStylesField = new CT_CellStyles();
            this.cellXfsField = new CT_CellXfs();
            this.cellStyleXfsField = new CT_CellStyleXfs();
            this.bordersField = new CT_Borders();
            this.fillsField = new CT_Fills();
            this.fontsField = new CT_Fonts();
            this.numFmtsField = new CT_NumFmts();
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
            this.extLstField = new CT_ExtensionList();
            this.protectionField = new CT_CellProtection();
            //this.borderField = new CT_Border();
            this.alignmentField = new CT_CellAlignment();
            //this.fillField = new CT_Fill();
            this.numFmtField = new CT_NumFmt();
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
    public class CT_Color
    {
        // all attributes are optional
        private bool? autoField = null;

        private uint? indexedField = null; // type of xsd:unsignedInt // TODO change all the uses theme to use uint instead of signed integer variants

        private byte[] rgbField = null; // TODO type of	ST_UnsignedIntHex

        private uint? themeField = null; // TODO change all the uses theme to use uint instead of signed integer variants

        private double? tintField = null;

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
        [XmlIgnore]
        // do not remove this field or change the name, because it is automatically used by the XmlSerializer to decide if the auto attribute should be printed or not.
        public bool autoSpecified
        {
            get { return (null != autoField); }
        }
        public bool IsSetAuto()
        {
            return autoSpecified;
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
            }
        }

        [XmlIgnore]
        public bool indexedSpecified
        {
            get { return (null != indexedField); }
        }
        public bool IsSetIndexed()
        {
            return indexedSpecified;
        }
        #endregion indexed

        #region rgb
        [XmlAttribute]
        // TODO this is Type ST_UnsignedIntHex 
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
        [XmlIgnore]
        public bool rgbSpecified
        {
            get { return (null != rgbField); }
        }
        public void SetRgb(byte R, byte G, byte B)
        {
            this.rgbField = new byte[3];
            this.rgbField[0] = R;
            this.rgbField[1] = G;
            this.rgbField[2] = B;
        }
        public bool IsSetRgb()
        {
            return rgbSpecified;
        }
        public void SetRgb(byte[] rgb)
        {
            Array.Copy(rgb, this.rgb, 3);
        }
        public byte[] GetRgb()
        {
            return rgb;
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
        [XmlIgnore]
        public bool themeSpecified
        {
            get { return (null != themeField); }

        }
        public bool IsSetTheme()
        {
            return themeSpecified;
        }
        #endregion theme

        #region tint
        [System.ComponentModel.DefaultValueAttribute(0.0D)]
        [XmlAttribute]
        public double tint
        {
            get
            {
                return (null == tintField) ? 0.0D : (double)this.tintField;
            }
            set
            {
                this.tintField = value;
            }
        }
        [XmlIgnore]
        public bool tintSpecified
        {
            get { return (null != tintField); }

        }
        public bool IsSetTint()
        {
            return tintSpecified;
        }

        #endregion tint
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_FontScheme
    {
        private ST_FontScheme valField;

        /// <remarks/>
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
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ElementName = "font")]
    public class CT_Font
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Font));
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });

        // all elements are optional
        private CT_FontName nameField = null; // name of the font
        private CT_IntProperty charsetField = null;
        private CT_IntProperty familyField = null; // family of the font
        private CT_BooleanProperty bField = null; // typeface bold
        private CT_BooleanProperty iField = null;   // italic
        private CT_BooleanProperty strikeField = null; //   strike through
        private CT_BooleanProperty outlineField = null;
        private CT_BooleanProperty shadowField = null;
        private CT_BooleanProperty condenseField = null;
        private CT_BooleanProperty extendField = null;
        private CT_Color colorField = null;
        private CT_FontSize szField = null; // size of the font
        private CT_UnderlineProperty uField = null; // underline
        private CT_VerticalAlignFontProperty vertAlignField = null;  // vertical alignment of the text
        private CT_FontScheme schemeField = null;


        public static CT_Font Parse(string xml)
        {
            CT_Font result;
            using (StringReader stream = new StringReader(xml))
            {
                result = (CT_Font)serializer.Deserialize(stream);
            }
            return result;
        }

        public static void Save(Stream stream, CT_Font font)
        {
            serializer.Serialize(stream, font, namespaces);
        }

        public static string GetString(CT_Font font)
        {
            using (StringWriter writer = new StringWriter())
            {
                serializer.Serialize(writer, font, namespaces);
                return writer.ToString();
            }
        }

        #region name
        [XmlElement]
        public CT_FontName name
        {
            get { return this.nameField; }
            set { this.nameField = value; }
        }
        [XmlIgnore]
        // do not remove this field or change the name, because it is automatically used by the XmlSerializer to decide if the name attribute should be printed or not.
        public bool nameSpecified
        {
            get { return (null != nameField); }
        }
        public int sizeOfNameArray()
        {
            return this.nameField == null ? 0 : 1;
        }
        public CT_FontName AddNewName()
        {
            this.nameField = new CT_FontName();
            return this.nameField;
        }
        public CT_FontName GetNameArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.nameField;
        }
        #endregion name

        #region charset
        [XmlElement]
        public CT_IntProperty charset
        {
            get { return this.charsetField; }
            set { this.charsetField = value; }
        }
        [XmlIgnore]
        public bool charsetSpecified
        {
            get { return (null != charsetField); }
        }
        public int sizeOfCharsetArray()
        {
            return this.charsetField == null ? 0 : 1;
        }
        public CT_IntProperty AddNewCharset()
        {
            this.charsetField = new CT_IntProperty();
            return this.charsetField;
        }
        public CT_IntProperty GetCharsetArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.charsetField;
        }
        #endregion charset

        #region family
        [XmlElement]
        public CT_IntProperty family
        {
            get { return this.familyField; }
            set { this.familyField = value; }
        }
        [XmlIgnore]
        public bool familySpecified
        {
            get { return (null != familyField); }
        }
        public int sizeOfFamilyArray()
        {
            return this.familyField == null ? 0 : 1;
        }
        public CT_IntProperty AddNewFamily()
        {
            this.familyField = new CT_IntProperty();
            return this.familyField;
        }
        //public void SetFamilyArray()
        //{
        //    this.familyField = null;
        //}
        public CT_IntProperty GetFamilyArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.familyField;
        }
        #endregion family

        #region b
        [XmlElement]
        public CT_BooleanProperty b
        {
            get { return this.bField; }
            set { this.bField = value; }
        }
        [XmlIgnore]
        public bool bSpecified
        {
            get { return (null != bField); }
        }
        public int sizeOfBArray()
        {
            return this.bField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewB()
        {
            this.bField = new CT_BooleanProperty();
            return this.bField;
        }
        public void SetBArray(CT_BooleanProperty[] array)
        {
            this.bField = array.Length > 0 ? array[0] : null;
        }
        public CT_BooleanProperty GetBArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.bField;
        }
        #endregion b

        #region i
        [XmlElement]
        public CT_BooleanProperty i
        {
            get { return this.iField; }
            set { this.iField = value; }
        }
        [XmlIgnore]
        public bool iSpecified
        {
            get { return (null != iField); }
        }
        public int sizeOfIArray()
        {
            return this.iField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewI()
        {
            this.iField = new CT_BooleanProperty();
            return this.iField;
        }
        public void SetIArray(CT_BooleanProperty[] array)
        {
            this.iField = array.Length > 0 ? array[0] : null;
        }
        public CT_BooleanProperty GetIArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.iField;
        }
        #endregion i

        #region strike
        [XmlElement]
        public CT_BooleanProperty strike
        {
            get { return this.strikeField; }
            set { this.strikeField = value; }
        }
        [XmlIgnore]
        public bool strikeSpecified
        {
            get { return (null != strikeField); }
        }
        public int sizeOfStrikeArray()
        {
            return this.strikeField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewStrike()
        {
            this.strikeField = new CT_BooleanProperty();
            return this.strikeField;
        }
        public void SetStrikeArray(CT_BooleanProperty[] array)
        {
            this.strikeField = array.Length > 0 ? array[0] : null;
        }
        public CT_BooleanProperty GetStrikeArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.strikeField;
        }
        #endregion strike

        #region outline
        [XmlElement]
        public CT_BooleanProperty outline
        {
            get { return this.outlineField; }
            set { this.outlineField = value; }
        }
        [XmlIgnore]
        public bool outlineSpecified
        {
            get { return (null != outlineField); }
        }
        public int sizeOfOutlineArray()
        {
            return this.outlineField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewOutline()
        {
            this.outlineField = new CT_BooleanProperty();
            return this.outlineField;
        }
        public void SetOutlineArray(CT_BooleanProperty[] array)
        {
            this.outlineField = array.Length > 0 ? array[0] : null;
        }
        public CT_BooleanProperty GetOutlineArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.outlineField;
        }
        #endregion outline

        #region shadow
        [XmlElement]
        public CT_BooleanProperty shadow
        {
            get { return this.shadowField; }
            set { this.shadowField = value; }
        }
        [XmlIgnore]
        public bool shadowSpecified
        {
            get { return (null != shadowField); }
        }
        public int sizeOfShadowArray()
        {
            return this.shadowField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewShadow()
        {
            this.shadowField = new CT_BooleanProperty();
            return this.shadowField;
        }
        public CT_BooleanProperty GetShadowArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.shadowField;
        }
        #endregion shadow

        #region condense
        [XmlElement]
        public CT_BooleanProperty condense
        {
            get { return this.condenseField; }
            set { this.condenseField = value; }
        }
        [XmlIgnore]
        public bool condenseSpecified
        {
            get { return (null != condenseField); }
        }
        public int sizeOfCondenseArray()
        {
            return this.condenseField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewCondense()
        {
            this.condenseField = new CT_BooleanProperty();
            return this.condenseField;
        }
        public CT_BooleanProperty GetCondenseArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.condenseField;
        }
        #endregion condense

        #region extend
        [XmlElement]
        public CT_BooleanProperty extend
        {
            get { return this.extendField; }
            set { this.extendField = value; }
        }
        [XmlIgnore]
        public bool extendSpecified
        {
            get { return (null != extendField); }
        }
        public int sizeOfExtendArray()
        {
            return this.extendField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewExtend()
        {
            this.extendField = new CT_BooleanProperty();
            return this.extendField;
        }
        public CT_BooleanProperty GetExtendArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.extendField;
        }
        #endregion extend

        #region color
        [XmlElement]
        public CT_Color color
        {
            get { return this.colorField; }
            set { this.colorField = value; }
        }
        [XmlIgnore]
        public bool colorSpecified
        {
            get { return (null != colorField); }
        }
        public int sizeOfColorArray()
        {
            return this.colorField == null ? 0 : 1;
        }
        public CT_Color GetColorArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.colorField;
        }
        public void SetColorArray(CT_Color[] array)
        {
            this.colorField = array.Length > 0 ? array[0] : null;
        }
        public CT_Color AddNewColor()
        {
            this.colorField = new CT_Color();
            return this.colorField;
        }
        #endregion color

        #region sz
        [XmlElement]
        public CT_FontSize sz
        {
            get { return this.szField; }
            set { this.szField = value; }
        }
        [XmlIgnore]
        public bool szSpecified
        {
            get { return (null != szField); }            
        }
        public int sizeOfSzArray()
        {
            return this.szField == null ? 0 : 1;
        }
        public CT_FontSize AddNewSz()
        {
            this.szField = new CT_FontSize();
            return this.szField;
        }
        public void SetSzArray(CT_FontSize[] array)
        {
            this.szField = array.Length > 0 ? array[0] : null;
        }
        public CT_FontSize GetSzArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.szField;
        }
        #endregion sz

        #region u
        [XmlElement]
        public CT_UnderlineProperty u
        {
            get { return this.uField; }
            set { this.uField = value; }
        }
        [XmlIgnore]
        public bool uSpecified
        {
            get { return (null != uField); }
        }
        public int sizeOfUArray()
        {
            return this.uField == null ? 0 : 1;
        }
        public CT_UnderlineProperty AddNewU()
        {
            this.uField = new CT_UnderlineProperty();
            return this.uField;
        }
        public void SetUArray(CT_UnderlineProperty[] array)
        {
            this.uField = array.Length > 0 ? array[0] : null;
        }
        public CT_UnderlineProperty GetUArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.uField;
        }
        #endregion u

        #region vertAlign
        [XmlElement]
        public CT_VerticalAlignFontProperty vertAlign
        {
            get { return this.vertAlignField; }
            set { this.vertAlignField = value; }
        }
        [XmlIgnore]
        public bool vertAlignSpecified
        {
            get { return (null != vertAlignField); }
        }
        public int sizeOfVertAlignArray()
        {
            return this.vertAlignField == null ? 0 : 1;
        }
        public CT_VerticalAlignFontProperty AddNewVertAlign()
        {
            this.vertAlignField = new CT_VerticalAlignFontProperty();
            return this.vertAlignField;
        }
        public void SetVertAlignArray(CT_VerticalAlignFontProperty[] array)
        {
            this.vertAlignField = array.Length > 0 ? array[0] : null;
        }
        public CT_VerticalAlignFontProperty GetVertAlignArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.vertAlignField;
        }
        #endregion vertAlign

        #region scheme
        [XmlElement]
        public CT_FontScheme scheme
        {
            get { return this.schemeField; }
            set { this.schemeField = value; }
        }
        [XmlIgnore]
        public bool schemeSpecified
        {
            get { return (null != schemeField); }
        }
        public int sizeOfSchemeArray()
        {
            return this.schemeField == null ? 0 : 1;
        }
        public CT_FontScheme AddNewScheme()
        {
            this.schemeField = new CT_FontScheme();
            return this.schemeField;
        }
        public CT_FontScheme GetSchemeArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.schemeField;
        }
        #endregion scheme
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_FontName
    {

        private string valField;

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
        /// <remarks/>
        none = 1,

        /// <remarks/>
        major = 2,

        /// <remarks/>
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

        [System.ComponentModel.DefaultValueAttribute(ST_UnderlineValues.single)]
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
        /// <remarks/>
        none,

        /// <remarks/>
        single,

        /// <remarks/>
        [XmlEnum("double")]
        @double,

        /// <remarks/>
        singleAccounting,

        /// <remarks/>
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
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_VerticalAlignRun
    {
        /// <remarks/>
        baseline,

        /// <remarks/>
        superscript,

        /// <remarks/>
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

        [System.ComponentModel.DefaultValueAttribute(true)]
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
    }


    public enum ST_TableStyleType
    {

        /// <remarks/>
        wholeTable,

        /// <remarks/>
        headerRow,

        /// <remarks/>
        totalRow,

        /// <remarks/>
        firstColumn,

        /// <remarks/>
        lastColumn,

        /// <remarks/>
        firstRowStripe,

        /// <remarks/>
        secondRowStripe,

        /// <remarks/>
        firstColumnStripe,

        /// <remarks/>
        secondColumnStripe,

        /// <remarks/>
        firstHeaderCell,

        /// <remarks/>
        lastHeaderCell,

        /// <remarks/>
        firstTotalCell,

        /// <remarks/>
        lastTotalCell,

        /// <remarks/>
        firstSubtotalColumn,

        /// <remarks/>
        secondSubtotalColumn,

        /// <remarks/>
        thirdSubtotalColumn,

        /// <remarks/>
        firstSubtotalRow,

        /// <remarks/>
        secondSubtotalRow,

        /// <remarks/>
        thirdSubtotalRow,

        /// <remarks/>
        blankRow,

        /// <remarks/>
        firstColumnSubheading,

        /// <remarks/>
        secondColumnSubheading,

        /// <remarks/>
        thirdColumnSubheading,

        /// <remarks/>
        firstRowSubheading,

        /// <remarks/>
        secondRowSubheading,

        /// <remarks/>
        thirdRowSubheading,

        /// <remarks/>
        pageFieldLabels,

        /// <remarks/>
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

        [System.ComponentModel.DefaultValueAttribute(true)]
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

        [System.ComponentModel.DefaultValueAttribute(true)]
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

        [System.Xml.Serialization.XmlIgnore]
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

        [System.Xml.Serialization.XmlIgnore]
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

        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1")]
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

        [System.Xml.Serialization.XmlIgnore]
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

        [System.Xml.Serialization.XmlIgnore]
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
