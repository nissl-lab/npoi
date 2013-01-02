
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(ElementName = "xf", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class CT_Xf
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Xf));
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });

        private CT_CellAlignment alignmentField = null;

        private CT_CellProtection protectionField = null;

        private CT_ExtensionList extLstField = null;

        private uint numFmtIdField = 0;
        private uint fontIdField = 0;
        private uint fillIdField = 0;
        private uint borderIdField = 0;
        private uint xfIdField = 0;
        private bool quotePrefixField = false;
        private bool pivotButtonField = false;
        private bool applyNumberFormatField = false;
        private bool applyFontField = false;
        private bool applyFillField = false;
        private bool applyBorderField = false;
        private bool applyAlignmentField = false;
        private bool applyProtectionField = false;

        bool numFmtIdSpecifiedField = false;
        bool fontIdSpecifiedField = false;
        bool fillIdSpecifiedField = false;
        bool borderIdSpecifiedField = false;
        bool xfIdSpecifiedField = false;

        public CT_Xf Copy()
        {
            CT_Xf obj = new CT_Xf();
            obj.alignmentField = this.alignmentField;
            obj.protectionField = this.protectionField;
            obj.extLstField = null == extLstField ? null : this.extLstField.Copy();

            obj.applyAlignmentField = this.applyAlignmentField;
            obj.applyBorderField = this.applyBorderField;
            obj.applyFillField = this.applyFillField;
            obj.applyFontField = this.applyFontField;
            obj.applyNumberFormatField = this.applyNumberFormatField;
            obj.applyProtectionField = this.applyProtectionField;
            obj.borderIdField = this.borderIdField;
            obj.fillIdField = this.fillIdField;
            obj.fontIdField = this.fontIdField;
            obj.numFmtIdField = this.numFmtIdField;
            obj.pivotButtonField = this.pivotButtonField;
            obj.quotePrefixField = this.quotePrefixField;
            obj.xfIdField = this.xfIdField;
            return obj;
        }

        public static CT_Xf Parse(string xml)
        {
            CT_Xf result;
            using (StringReader stream = new StringReader(xml))
            {
                result = (CT_Xf)serializer.Deserialize(stream);
            }
            return result;
        }
        public static void Save(Stream stream, CT_Xf font)
        {
            serializer.Serialize(stream, font, namespaces);
        }
        public override string ToString()
        {
            using (StringWriter stream = new StringWriter())
            {
                serializer.Serialize(stream, this, namespaces);
                return stream.ToString();
            }
        }

        public bool IsSetFontId()
        {
            return this.fontIdSpecifiedField;
        }
        public bool IsSetAlignment()
        {
            return alignmentField!=null;
        }
        public void UnsetAlignment()
        {
            this.alignmentField = null;
        }

        public bool IsSetExtLst()
        {
            return this.extLst == null;
        }
        public void UnsetExtLst()
        {
            this.extLst = null;
        }

        public bool IsSetProtection()
        {
            return this.protection!=null;
        }
        public void UnsetProtection()
        {
            this.protection = null;
        }
        public bool IsSetLocked()
        {
            // first guess:
            return IsSetProtection() && protectionField.lockedSpecified && (protectionField.locked == true);
        }
        public CT_CellProtection AddNewProtection()
        {
            this.protectionField = new CT_CellProtection();
            return this.protectionField;
        }

        [XmlElement]
        public CT_CellAlignment alignment
        {
            get { return this.alignmentField; }
            set { this.alignmentField = value; }
        }


        [XmlElement]
        public CT_CellProtection protection
        {
            get { return this.protectionField; }
            set { this.protectionField = value; }
        }

        [XmlElement]
        public CT_ExtensionList extLst
        {
            get { return this.extLstField; }
            set { this.extLstField = value; }
        }
        [XmlAttribute]
        public uint numFmtId
        {
            get { return this.numFmtIdField; }
            set { 
                this.numFmtIdField = value;
                this.numFmtIdSpecifiedField = true;
            }
        }

        [XmlIgnore]
        public bool numFmtIdSpecified
        {
            get { return numFmtIdSpecifiedField; }
            set { numFmtIdSpecifiedField=value; }
        }

        [XmlAttribute]
        public uint fontId
        {
            get { return this.fontIdField; }
            set { 
                this.fontIdField = value;
                this.fontIdSpecifiedField = true;
            }
        }

        [XmlIgnore]
        public bool fontIdSpecified
        {
            get { return fontIdSpecifiedField; }
            set { fontIdSpecifiedField = value; }
        }

        [XmlAttribute]
        public uint fillId
        {
            get { return this.fillIdField; }
            set { 
                this.fillIdField = value;
                this.fillIdSpecifiedField = true;
            }
        }

        [XmlIgnore]
        public bool fillIdSpecified
        {
            get { return fillIdSpecifiedField; }
            set { fillIdSpecifiedField = value; }
        }

        [XmlAttribute]
        public uint borderId
        {
            get { return this.borderIdField; }
            set { 
                this.borderIdField = value;
                borderIdSpecifiedField = true;
            }
        }
        [XmlIgnore]
        public bool borderIdSpecified
        {
            get { return borderIdSpecifiedField; }
            set { borderIdSpecifiedField = value; }
        }

        [XmlAttribute]
        public uint xfId
        {
            get { return this.xfIdField; }
            set { 
                this.xfIdField = value;
                this.xfIdSpecifiedField = true;
            }
        }
        [XmlIgnore]
        public bool xfIdSpecified
        {
            get { return xfIdSpecifiedField; }
            set { xfIdSpecifiedField = value; }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool quotePrefix
        {
            get { return quotePrefixField; }
            set { this.quotePrefixField = value; }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool pivotButton
        {
            get { return pivotButtonField; }
            set { this.pivotButtonField = value; }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyNumberFormat
        {
            get { return this.applyNumberFormatField; }
            set
            {
                this.applyNumberFormatField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyFont
        {
            get { return this.applyFontField; }
            set { this.applyFontField = value; }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyFill
        {
            get { return this.applyFillField; }
            set { this.applyFillField = value; }
        }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyBorder
        {
            get { return this.applyBorderField; }
            set { this.applyBorderField = value; }
        }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyAlignment
        {
            get { return this.applyAlignmentField; }
            set { this.applyAlignmentField = value; }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyProtection
        {
            get { return this.applyProtectionField; }
            set { this.applyProtectionField = value; }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_CellStyleXfs
    {

        private List<CT_Xf> xfField = null;

        private uint countField;

        private bool countFieldSpecified;

        public CT_CellStyleXfs()
        {
            //this.xfField = new List<CT_Xf>();
        }
        public CT_Xf AddNewXf()
        {
            if (this.xfField == null)
                this.xfField = new List<CT_Xf>();
            CT_Xf xf = new CT_Xf();
            this.xfField.Add(xf);
            return xf;
        }
        [XmlElement]
        public List<CT_Xf> xf
        {
            get
            {
                return this.xfField;
            }
            set
            {
                this.xfField = value;
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

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Border
    {

        private CT_BorderPr leftField;

        private CT_BorderPr rightField;

        private CT_BorderPr topField;

        private CT_BorderPr bottomField;

        private CT_BorderPr diagonalField;

        private CT_BorderPr verticalField;

        private CT_BorderPr horizontalField;

        private bool diagonalUpField;

        private bool diagonalUpFieldSpecified;

        private bool diagonalDownField;

        private bool diagonalDownFieldSpecified;

        private bool outlineField;

        public CT_Border()
        {
            //this.horizontalField = new CT_BorderPr();
            //this.verticalField = new CT_BorderPr();
            //this.diagonalField = new CT_BorderPr();
            //this.bottomField = new CT_BorderPr();
            //this.topField = new CT_BorderPr();
            //this.rightField = new CT_BorderPr();
            //this.leftField = new CT_BorderPr();
            //this.outlineField = true;
        }
        public CT_Border Copy()
        {
            CT_Border obj = new CT_Border();
            obj.bottomField = this.bottomField;
            obj.topField = this.topField;
            obj.rightField = this.rightField;
            obj.leftField = this.leftField;
            obj.horizontalField = this.horizontalField;
            obj.verticalField = this.verticalField;
            obj.outlineField = this.outlineField;
            return obj;
        }
        public CT_BorderPr AddNewDiagonal()
        {
            if(this.diagonalField==null)
                this.diagonalField = new CT_BorderPr();
            return this.diagonalField;
        }
        public bool IsSetDiagonal()
        {
            return this.diagonalField != null;
        }
        public void unsetDiagonal()
        {
            this.diagonalField = null;
        }

        public void unsetRight()
        {
            this.rightField = null;
        }
        public void unsetLeft()
        {
            this.leftField = null;
        }
        public void unsetTop()
        {
            this.topField = null;
        }
        public void unsetBottom()
        {
            this.bottomField = null;
        }
        public bool IsSetBottom()
        {
            return this.bottomField != null;
        }
        public bool IsSetLeft()
        {
            return this.leftField != null;
        }
        public bool IsSetRight()
        {
            return this.rightField != null;
        }
        public bool IsSetTop()
        {
            return this.topField != null;
        }

        public bool IsSetBorder()
        {
            return this.leftField != null || this.rightField != null
                || this.topField != null || this.bottomField != null;
        }
        public CT_BorderPr AddNewTop()
        {
            if (this.topField == null)
                this.topField = new CT_BorderPr();
            return this.topField;
        }
        public CT_BorderPr AddNewRight()
        {
            if (this.rightField == null)
                this.rightField = new CT_BorderPr();
            return this.rightField;
        }
        public CT_BorderPr AddNewLeft()
        {
            if (this.leftField == null)
                this.leftField = new CT_BorderPr();
            return this.leftField;
        }
        public CT_BorderPr AddNewBottom()
        {
            if (this.bottomField == null)
                this.bottomField = new CT_BorderPr();
            return this.bottomField;
        }

        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr left
        {
            get
            {
                return this.leftField;
            }
            set
            {
                this.leftField = value;
            }
        }
        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr right
        {
            get
            {
                return this.rightField;
            }
            set
            {
                this.rightField = value;
            }
        }
        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr top
        {
            get
            {
                return this.topField;
            }
            set
            {
                this.topField = value;
            }
        }
        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr bottom
        {
            get
            {
                return this.bottomField;
            }
            set
            {
                this.bottomField = value;
            }
        }
        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr diagonal
        {
            get
            {
                return this.diagonalField;
            }
            set
            {
                this.diagonalField = value;
            }
        }
        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr vertical
        {
            get
            {
                return this.verticalField;
            }
            set
            {
                this.verticalField = value;
            }
        }
        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr horizontal
        {
            get
            {
                return this.horizontalField;
            }
            set
            {
                this.horizontalField = value;
            }
        }
        [XmlAttribute]
        public bool diagonalUp
        {
            get
            {
                return this.diagonalUpField;
            }
            set
            {
                this.diagonalUpField = value;
            }
        }

        [XmlIgnore]
        public bool diagonalUpSpecified
        {
            get
            {
                return this.diagonalUpFieldSpecified;
            }
            set
            {
                this.diagonalUpFieldSpecified = value;
            }
        }
        [XmlAttribute]
        public bool diagonalDown
        {
            get
            {
                return this.diagonalDownField;
            }
            set
            {
                this.diagonalDownField = value;
            }
        }

        [XmlIgnore]
        public bool diagonalDownSpecified
        {
            get
            {
                return this.diagonalDownFieldSpecified;
            }
            set
            {
                this.diagonalDownFieldSpecified = value;
            }
        }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool outline
        {
            get
            {
                return this.outlineField;
            }
            set
            {
                this.outlineField = value;
            }
        }
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Border));
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });
        public override string ToString()
        {
            StringWriter stringWriter = new StringWriter();
            serializer.Serialize(stringWriter, this);
            return stringWriter.ToString();
        }
    }

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

    public enum ST_BorderStyle
    {
        none,
        thin,
        medium,
        dashed,
        dotted,
        thick,
        @double,
        hair,
        mediumDashed,
        dashDot,
        mediumDashDot,
        dashDotDot,
        mediumDashDotDot,
        slantDashDot,
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Borders
    {
        private List<CT_Border> borderField;
        private uint countField = 0;
        private bool countFieldSpecified = false;

        public CT_Borders()
        {
            this.borderField = new List<CT_Border>();
        }
        public CT_Border AddNewBorder()
        {
            CT_Border border = new CT_Border();
            this.borderField.Add(border);
            return border;
        }
        [XmlElement]
        public List<CT_Border> border
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
        public void SetBorderArray(List<CT_Border> array)
        {
            borderField = array;
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

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Fills
    {
        private List<CT_Fill> fillField;
        private uint countField = 0;
        private bool countFieldSpecified = false;

        public CT_Fills()
        {
            this.fillField = new List<CT_Fill>();
        }
        [XmlElement]
        public List<CT_Fill> fill
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
        public void SetFillArray(List<CT_Fill> array)
        {
            fillField = array;
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

    //[Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(ElementName = "fonts", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
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
            if (array != null)
                fontField = new List<CT_Font>(array);
            else
                fontField.Clear();
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
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Fonts));
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });
        public override string ToString()
        {
            StringWriter stringWriter = new StringWriter();
            serializer.Serialize(stringWriter, this, namespaces);
            return stringWriter.ToString();
        }
    }
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
    public class CT_Fill
    {
        private CT_PatternFill patternFillField = null;
        private CT_GradientFill gradientFillField = null;

        [XmlElement]
        public CT_PatternFill patternFill
        {
            get            {                return this.patternFillField;            }
            set            {                this.patternFillField = value;            }
        }
        [XmlElement]
        public CT_GradientFill gradientFill
        {
            get { return this.gradientFillField; }
            set { this.gradientFillField = value; }
        }

        public CT_PatternFill GetPatternFill()
        {
            return this.patternFillField;
        }

        public CT_PatternFill AddNewPatternFill()
        {
            this.patternFillField = new CT_PatternFill();
            return GetPatternFill();
        }
        public bool IsSetPatternFill()
        {
            return this.patternFillField != null;
        }
        public CT_Fill Copy()
        {
            CT_Fill obj = new CT_Fill();
            obj.patternFillField = this.patternFillField;
            return obj;
        }
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Fill));
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });

        public override string ToString()
        {
            StringWriter stringWriter = new StringWriter();
            serializer.Serialize(stringWriter, this, namespaces);
            return stringWriter.ToString();
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
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });
        public override string ToString()
        {
            using (StringWriter stringWriter = new StringWriter())
            {
                serializer.Serialize(stringWriter, this, namespaces);
                return stringWriter.ToString();
            }
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
        private List<CT_FontName> nameField = null; // name of the font
        private List<CT_IntProperty> charsetField = null;
        private List<CT_IntProperty> familyField = null; // family of the font
        private List<CT_BooleanProperty> bField = null; // typeface bold
        private List<CT_BooleanProperty> iField = null;   // italic
        private List<CT_BooleanProperty> strikeField = null; //   strike through
        private CT_BooleanProperty outlineField = null;
        private CT_BooleanProperty shadowField = null;
        private CT_BooleanProperty condenseField = null;
        private CT_BooleanProperty extendField = null;
        private List<CT_Color> colorField = null;
        private List<CT_FontSize> szField = null; // size of the font
        private List<CT_UnderlineProperty> uField = null; // underline
        private List<CT_VerticalAlignFontProperty> vertAlignField = null;  // vertical alignment of the text
        private List<CT_FontScheme> schemeField = null;

        public CT_Font()
        {
            this.nameField = new List<CT_FontName>();
            this.szField = new List<CT_FontSize>();
            this.colorField = new List<CT_Color>();
            this.familyField = new List<CT_IntProperty>();
            this.charsetField = new List<CT_IntProperty>();
            this.uField = new List<CT_UnderlineProperty>();
            this.bField = new List<CT_BooleanProperty>();
            this.iField = new List<CT_BooleanProperty>();
            this.vertAlignField = new List<CT_VerticalAlignFontProperty>();
            this.schemeField = new List<CT_FontScheme>();
            this.strikeField = new List<CT_BooleanProperty>();
        }

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
        public List<CT_FontName> name
        {
            get { return this.nameField; }
            set { this.nameField = value; }
        }
        public int sizeOfNameArray()
        {
            return this.nameField.Count;
        }
        public CT_FontName AddNewName()
        {       
            CT_FontName fn = new CT_FontName();
            this.nameField.Add(fn);
            return fn;
        }
        public void SetNameArray(int index, CT_FontName value)
        {
            this.nameField[index]= value;
        }
        public CT_FontName GetNameArray(int index)
        {
            return this.nameField[index];
        }
        #endregion name

        #region charset
        [XmlElement]
        public List<CT_IntProperty> charset
        {
            get { return this.charsetField; }
            set { this.charsetField = value; }
        }
        public int sizeOfCharsetArray()
        {
            return this.charsetField.Count;
        }
        public CT_IntProperty AddNewCharset()
        {
            CT_IntProperty prop = new CT_IntProperty();
            this.charsetField.Add(prop);
            return prop;
        }
        public void SetCharsetArray(int index, CT_IntProperty value)
        {
            this.charsetField[index]= value;
        }
        public CT_IntProperty GetCharsetArray(int index)
        {
            return this.charsetField[index];
        }
        #endregion charset

        #region family
        [XmlElement]
        public List<CT_IntProperty> family
        {
            get { return this.familyField; }
            set { this.familyField = value; }
        }
        public int sizeOfFamilyArray()
        {
            return this.familyField.Count;
        }
        public CT_IntProperty AddNewFamily()
        {
            CT_IntProperty newfamily = new CT_IntProperty();
            this.familyField.Add(newfamily);
            return newfamily;
        }
        public void SetFamilyArray(int index, CT_IntProperty value)
        {
            this.familyField[index]=value;
        }
        public CT_IntProperty GetFamilyArray(int index)
        {
            return this.familyField[index];
        }
        #endregion family

        #region b
        [XmlElement]
        public List<CT_BooleanProperty> b
        {
            get { return this.bField; }
            set { this.bField = value; }
        }
        public int sizeOfBArray()
        {
            return this.bField.Count;
        }
        public CT_BooleanProperty AddNewB()
        {
            CT_BooleanProperty newB = new CT_BooleanProperty();
            this.bField.Add(newB);
            return newB;
        }
        public void SetBArray(int index, CT_BooleanProperty value)
        {
            this.bField[index] = value;
        }
        public void SetBArray(CT_BooleanProperty[] array)
        {
            if (array == null)
                this.bField = null;
            else
                this.bField = new List<CT_BooleanProperty>(array);
        }
        public CT_BooleanProperty GetBArray(int index)
        {
            return this.bField[index];
        }
        #endregion b

        #region i
        [XmlElement]
        public List<CT_BooleanProperty> i
        {
            get { return this.iField; }
            set { this.iField = value; }
        }
        public int sizeOfIArray()
        {
            return this.iField.Count;
        }
        public CT_BooleanProperty AddNewI()
        {
            CT_BooleanProperty newI=new CT_BooleanProperty();
            this.iField.Add(newI);
            return newI;
        }
        public void SetIArray(int index, CT_BooleanProperty value)
        {
            this.iField[index]= value;
        }
        public void SetIArray(CT_BooleanProperty[] array)
        {
            if (array == null)
                this.iField = new List<CT_BooleanProperty>();
            else
                this.iField = new List<CT_BooleanProperty>(array);
        }
        public CT_BooleanProperty GetIArray(int index)
        {
            return this.iField[index];
        }
        #endregion i

        #region strike
        [XmlElement]
        public List<CT_BooleanProperty> strike
        {
            get { return this.strikeField; }
            set { this.strikeField = value; }
        }
        public int sizeOfStrikeArray()
        {
            return this.strikeField.Count;
        }
        public CT_BooleanProperty AddNewStrike()
        {
            CT_BooleanProperty prop = new CT_BooleanProperty();
            this.strikeField.Add(prop);
            return prop;
        }
        public void SetStrikeArray(int index, CT_BooleanProperty value)
        {
            this.strikeField[index]= value;
        }
        public void SetStrikeArray(CT_BooleanProperty[] array)
        {
            this.strikeField = new List<CT_BooleanProperty>(array);
        }
        public CT_BooleanProperty GetStrikeArray(int index)
        {
            return this.strikeField[index];
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
        public List<CT_Color> color
        {
            get { return this.colorField; }
            set { this.colorField = value; }
        }
        public int sizeOfColorArray()
        {
            return this.colorField.Count;
        }
        public CT_Color GetColorArray(int index)
        {
            return this.colorField[index];
        }
        public void SetColorArray(int index, CT_Color value)
        {
            this.colorField[index]=value;
        }
        public void SetColorArray(CT_Color[] array)
        {
            this.colorField = new List<CT_Color>(array);
        }
        public CT_Color AddNewColor()
        {
            CT_Color newColor = new CT_Color();
            this.colorField.Add(newColor);
            return newColor;
        }
        #endregion color

        #region sz
        [XmlElement]
        public List<CT_FontSize> sz
        {
            get { return this.szField; }
            set { this.szField = value; }
        }
        public int sizeOfSzArray()
        {
            return this.szField.Count;
        }
        public CT_FontSize AddNewSz()
        {
            CT_FontSize newFs = new CT_FontSize();
            this.szField.Add(newFs);
            return newFs;
        }
        public void SetSzArray(int index, CT_FontSize value)
        {
            this.szField[index]= value;
        }
        public void SetSzArray(CT_FontSize[] array)
        {
            this.szField = new List<CT_FontSize>(array);
        }
        public CT_FontSize GetSzArray(int index)
        {
            return this.szField[index];
        }
        #endregion sz

        #region u
        [XmlElement]
        public List<CT_UnderlineProperty> u
        {
            get { return this.uField; }
            set { this.uField = value; }
        }
        public int sizeOfUArray()
        {
            return this.uField.Count;
        }
        public CT_UnderlineProperty AddNewU()
        {
            CT_UnderlineProperty newU = new CT_UnderlineProperty();
            this.uField.Add(newU);
            return newU;
        }
        public void SetUArray(int index,CT_UnderlineProperty value)
        {
            if (uField[index] != null)
                this.uField[index] = value;
            else
                this.uField.Insert(index, value);
        }
        public void SetUArray(CT_UnderlineProperty[] array)
        {
            this.uField = new List<CT_UnderlineProperty>(array);
        }
        public CT_UnderlineProperty GetUArray(int index)
        {
            return this.uField[index];
        }
        #endregion u

        #region vertAlign
        [XmlElement]
        public List<CT_VerticalAlignFontProperty> vertAlign
        {
            get { return this.vertAlignField; }
            set { this.vertAlignField = value; }
        }
        public int sizeOfVertAlignArray()
        {
            return this.vertAlignField.Count;
        }
        public CT_VerticalAlignFontProperty AddNewVertAlign()
        {
            CT_VerticalAlignFontProperty prop = new CT_VerticalAlignFontProperty();
            this.vertAlignField.Add(prop);
            return prop;
        }
        public void SetVertAlignArray(int index, CT_VerticalAlignFontProperty value)
        {
            this.vertAlignField[index]= value;
        }
        public void SetVertAlignArray(CT_VerticalAlignFontProperty[] array)
        {
            this.vertAlignField = new List<CT_VerticalAlignFontProperty>(array);
        }
        public CT_VerticalAlignFontProperty GetVertAlignArray(int index)
        {
            return this.vertAlignField[index];
        }
        #endregion vertAlign

        #region scheme
        [XmlElement]
        public List<CT_FontScheme> scheme
        {
            get { return this.schemeField; }
            set { this.schemeField = value; }
        }
        public int sizeOfSchemeArray()
        {
            return this.schemeField.Count;
        }
        public CT_FontScheme AddNewScheme()
        {
            CT_FontScheme newScheme = new CT_FontScheme();
            this.schemeField.Add(newScheme);
            return newScheme;
        }
        public void SetSchemeArray(int index, CT_FontScheme value)
        {
            this.schemeField[index]= value;
        }
        public CT_FontScheme GetSchemeArray(int index)
        {
            return this.schemeField[index];
        }
        #endregion scheme

        public override string ToString()
        {
            return CT_Font.GetString(this);
        }
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
