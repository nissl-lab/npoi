using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
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
            throw new NotImplementedException();
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(ElementName="styleSheet",Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
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
    public class CT_Color
    {
        private bool? autoField;

        private bool autoFieldSpecified;

        private long? indexedField;

        private bool indexedFieldSpecified;

        private byte[] rgbField;

        private long? themeField;

        private bool themeFieldSpecified;

        private double? tintField;

        public CT_Color()
        {
            this.tintField = 0D;
        }
        public bool IsSetIndexed()
        {
            return this.indexedField != null;
        }
        public bool IsSetAuto()
        {
            return this.autoField != null;
        }
        public bool IsSetTheme()
        {
            return this.themeField != null;
        }
        public bool IsSetTint()
        {
            return this.tintField != null;
        }
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool autoSpecified
        {
            get
            {
                return this.autoFieldSpecified;
            }
            set
            {
                this.autoFieldSpecified = value;
            }
        }

        public long indexed
        {
            get
            {
                return (long)this.indexedField;
            }
            set
            {
                this.indexedField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool indexedSpecified
        {
            get
            {
                return this.indexedFieldSpecified;
            }
            set
            {
                this.indexedFieldSpecified = value;
            }
        }

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

        public long theme
        {
            get
            {
                return (long)this.themeField;
            }
            set
            {
                this.themeField = value;
            }
        }
        public void SetVal(long val)
        {
            throw new NotImplementedException();
        }
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool themeSpecified
        {
            get
            {
                return this.themeFieldSpecified;
            }
            set
            {
                this.themeFieldSpecified = value;
            }
        }

        [System.ComponentModel.DefaultValueAttribute(0D)]
        public double tint
        {
            get
            {
                return (double)this.tintField;
            }
            set
            {
                this.tintField = value;
            }
        }

        public void SetRgb(byte R, byte G, byte B)
        {
            
        }
        public bool IsSetRgb()
        {
            return rgb != null;
        }
        public void SetRgb(byte[] rgb)
        {
            Array.Copy(rgb, this.rgb, 3);
        }
        public byte[] GetRgb()
        {
            return rgb;
        }

    }
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public class CT_FontScheme
    {

        private ST_FontScheme valField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
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
    public enum FontElementNameType
    {
        b,
        charset,
        color,
        condense,
        extend,
        family,
        i,
        name,
        outline,
        rFont,
        scheme,
        shadow,
        strike,
        sz,
        u,
        vertAlign
    }
    public class CT_Font
    {

        private List<object> itemsField;

        private List<FontElementNameType> itemsElementNameField;

        public CT_Font()
        {
            this.itemsElementNameField = new List<FontElementNameType>();
            this.itemsField = new List<object>();
        }
        public static CT_Font Parse(string xml)
        {
            throw new NotImplementedException();
        }

        [System.Xml.Serialization.XmlElementAttribute("b", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("charset", typeof(CT_IntProperty))]
        [System.Xml.Serialization.XmlElementAttribute("color", typeof(CT_Color))]
        [System.Xml.Serialization.XmlElementAttribute("condense", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("extend", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("family", typeof(CT_IntProperty))]
        [System.Xml.Serialization.XmlElementAttribute("i", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("name", typeof(CT_FontName))]
        [System.Xml.Serialization.XmlElementAttribute("outline", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("scheme", typeof(CT_FontScheme))]
        [System.Xml.Serialization.XmlElementAttribute("shadow", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("strike", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("sz", typeof(CT_FontSize))]
        [System.Xml.Serialization.XmlElementAttribute("u", typeof(CT_UnderlineProperty))]
        [System.Xml.Serialization.XmlElementAttribute("vertAlign", typeof(CT_VerticalAlignFontProperty))]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public object[] Items
        {
            get
            {
                return this.itemsField.ToArray();
            }
            set
            {
                this.itemsField = new List<object>(value);
            }
        }

        [System.Xml.Serialization.XmlElementAttribute("ItemsElementName")]
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public FontElementNameType[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                this.itemsElementNameField = new List<FontElementNameType>(value);
            }
        }
        public int sizeOfFamilyArray()
        {
            throw new NotImplementedException();
        }
        public CT_IntProperty GetFamilyArray(int index)
        {
            throw new NotImplementedException();
        }
        public CT_IntProperty AddNewFamily()
        {
            throw new NotImplementedException();
        }
        public CT_FontName AddNewName()
        {
            throw new NotImplementedException();
        }
        public CT_IntProperty AddNewCharset()
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty AddNewB()
        {
            throw new NotImplementedException();
        }
        public CT_UnderlineProperty AddNewU()
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty AddNewCondense()
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty AddNewExtend()
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty AddNewOutline()
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty AddNewShadow()
        {
            throw new NotImplementedException();
        }
        public CT_VerticalAlignFontProperty AddNewVertAlign()
        {
            throw new NotImplementedException();
        }
        public CT_VerticalAlignFontProperty GetVertAlignArray(int index)
        {
            throw new NotImplementedException();
        }
        public void SetFamilyArray()
        {
            throw new NotImplementedException();
        }
        public int sizeOfBArray()
        {
            throw new NotImplementedException();
        }

        public int sizeOfCharSetArray()
        {
            throw new NotImplementedException();
        }

        public int sizeOfColorArray()
        {
            throw new NotImplementedException();
        }
        public CT_Color GetColorArray(int index)
        {
            throw new NotImplementedException();
        }
        public CT_IntProperty GetCharsetArray(int index)
        {
            throw new NotImplementedException();
        }
        public CT_FontScheme GetSchemeArray(int index)
        {
            throw new NotImplementedException();
        }
        public int sizeOfSzArray()
        {
            throw new NotImplementedException();
        }

        public int sizeOfIArray()
        {
            throw new NotImplementedException();
        }
        public int sizeOfStrikeArray()
        {
            throw new NotImplementedException();
        }
        public int sizeOfNameArray()
        {
            throw new NotImplementedException();
        }
        public int sizeOfSchemeArray()
        {
            throw new NotImplementedException();
        }
        public int sizeOfUArray()
        {
            throw new NotImplementedException();
        }
        public int sizeOfCharsetArray()
        {
            throw new NotImplementedException();
        }
        public int sizeOfCondenseArray()
        {
            throw new NotImplementedException();
        }
        public int sizeOfExtendArray()
        {
            throw new NotImplementedException();
        }
        public int sizeOfOutlineArray()
        {
            throw new NotImplementedException();
        }
        public int sizeOfShadowArray()
        {
            throw new NotImplementedException();
        }
        public void SetUArray(CT_UnderlineProperty[] array)
        {
            throw new NotImplementedException();
        }
        public CT_UnderlineProperty GetUArray(int index)
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty GetStrikeArray(int index)
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty GetCondenseArray(int index)
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty GetExtendArray(int index)
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty GetShadowArray(int index)
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty GetOutlineArray(int index)
        {
            throw new NotImplementedException();
        }

        public void SetStrikeArray(CT_BooleanProperty property)
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty AddNewStrike()
        {
            throw new NotImplementedException();
        }

        public CT_FontName GetNameArray(int index)
        {
            throw new NotImplementedException();
        }
        public void SetSzArray(CT_FontSize[] array)
        {
            throw new NotImplementedException();
        }
        public CT_FontSize GetSzArray(int index)
        {
            throw new NotImplementedException();
        }
        public int SetColorArray(CT_Color[] array)
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty GetBArray(int index)
        {
            throw new NotImplementedException();
        }
        public int SetFontArray(CT_Font[] font)
        {
            throw new NotImplementedException();
        }

        public int SetVertAlignArray(CT_VerticalAlignFontProperty[] array)
        {
            throw new NotImplementedException();
        }
        public int sizeOfVertAlignArray()
        {
            throw new NotImplementedException();
        }

        public CT_Color AddNewColor()
        {
            throw new NotImplementedException();
        }
        public CT_FontSize AddNewSz()
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty GetIArray(int index)
        {
            throw new NotImplementedException();
        }
        public CT_BooleanProperty AddNewI()
        {
            throw new NotImplementedException();
        }
        public void SetIArray(CT_BooleanProperty[] array)
        {
            throw new NotImplementedException();
        }
        public void SetBArray(object[] array)
        {
            throw new NotImplementedException();
        }
        public CT_FontScheme AddNewScheme()
        {
            throw new NotImplementedException();
        }
    }


    public class CT_FontName
    {

        private string valField;

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


    public class CT_FontSize
    {

        private double valField;

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


    public enum ST_FontScheme
    {

        /// <remarks/>
        none,

        /// <remarks/>
        major,

        /// <remarks/>
        minor,
    }


    public class CT_UnderlineProperty
    {

        private ST_UnderlineValues valField;

        public CT_UnderlineProperty()
        {
            this.valField = ST_UnderlineValues.single;
        }

        [System.ComponentModel.DefaultValueAttribute(ST_UnderlineValues.single)]
        public ST_UnderlineValues val
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

    public enum ST_UnderlineValues
    {

        /// <remarks/>
        single,

        /// <remarks/>
        @double,

        /// <remarks/>
        singleAccounting,

        /// <remarks/>
        doubleAccounting,

        /// <remarks/>
        none,
    }

    public class CT_VerticalAlignFontProperty
    {

        private ST_VerticalAlignRun valField;

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

    public enum ST_VerticalAlignRun
    {

        /// <remarks/>
        baseline,

        /// <remarks/>
        superscript,

        /// <remarks/>
        subscript,
    }
    public class CT_BooleanProperty
    {

        private bool valField;

        public CT_BooleanProperty()
        {
            this.valField = true;
        }

        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool val
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

    public class CT_IntProperty
    {

        private int valField;

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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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
