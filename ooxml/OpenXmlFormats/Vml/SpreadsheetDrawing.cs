using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Vml
{
    /// <remarks/>
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:excel")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:excel", IsNullable = true)]
    public partial class CT_ClientData
    {

        private object[] itemsField;

        private ItemsChoiceType[] itemsElementNameField;

        private ST_ObjectType objectTypeField;

        /// <remarks/>
        [XmlElement("Accel", typeof(string), DataType = "integer")]
        [XmlElement("Accel2", typeof(string), DataType = "integer")]
        [XmlElement("Anchor", typeof(string))]
        [XmlElement("AutoFill", typeof(ST_TrueFalseBlank))]
        [XmlElement("AutoLine", typeof(ST_TrueFalseBlank))]
        [XmlElement("AutoPict", typeof(ST_TrueFalseBlank))]
        [XmlElement("AutoScale", typeof(ST_TrueFalseBlank))]
        [XmlElement("CF", typeof(ST_CF))]
        [XmlElement("Camera", typeof(ST_TrueFalseBlank))]
        [XmlElement("Cancel", typeof(ST_TrueFalseBlank))]
        [XmlElement("Checked", typeof(string), DataType = "integer")]
        [XmlElement("ColHidden", typeof(ST_TrueFalseBlank))]
        [XmlElement("Colored", typeof(ST_TrueFalseBlank))]
        [XmlElement("Column", typeof(string), DataType = "integer")]
        [XmlElement("DDE", typeof(ST_TrueFalseBlank))]
        [XmlElement("Default", typeof(ST_TrueFalseBlank))]
        [XmlElement("DefaultSize", typeof(ST_TrueFalseBlank))]
        [XmlElement("Disabled", typeof(ST_TrueFalseBlank))]
        [XmlElement("Dismiss", typeof(ST_TrueFalseBlank))]
        [XmlElement("DropLines", typeof(string), DataType = "integer")]
        [XmlElement("DropStyle", typeof(string))]
        [XmlElement("Dx", typeof(string), DataType = "integer")]
        [XmlElement("FirstButton", typeof(ST_TrueFalseBlank))]
        [XmlElement("FmlaGroup", typeof(string))]
        [XmlElement("FmlaLink", typeof(string))]
        [XmlElement("FmlaMacro", typeof(string))]
        [XmlElement("FmlaPict", typeof(string))]
        [XmlElement("FmlaRange", typeof(string))]
        [XmlElement("FmlaTxbx", typeof(string))]
        [XmlElement("Help", typeof(ST_TrueFalseBlank))]
        [XmlElement("Horiz", typeof(ST_TrueFalseBlank))]
        [XmlElement("Inc", typeof(string), DataType = "integer")]
        [XmlElement("JustLastX", typeof(ST_TrueFalseBlank))]
        [XmlElement("LCT", typeof(string))]
        [XmlElement("ListItem", typeof(string))]
        [XmlElement("LockText", typeof(ST_TrueFalseBlank))]
        [XmlElement("Locked", typeof(ST_TrueFalseBlank))]
        [XmlElement("MapOCX", typeof(ST_TrueFalseBlank))]
        [XmlElement("Max", typeof(string), DataType = "integer")]
        [XmlElement("Min", typeof(string), DataType = "integer")]
        [XmlElement("MoveWithCells", typeof(ST_TrueFalseBlank))]
        [XmlElement("MultiLine", typeof(ST_TrueFalseBlank))]
        [XmlElement("MultiSel", typeof(string))]
        [XmlElement("NoThreeD", typeof(ST_TrueFalseBlank))]
        [XmlElement("NoThreeD2", typeof(ST_TrueFalseBlank))]
        [XmlElement("Page", typeof(string), DataType = "integer")]
        [XmlElement("PrintObject", typeof(ST_TrueFalseBlank))]
        [XmlElement("RecalcAlways", typeof(ST_TrueFalseBlank))]
        [XmlElement("Row", typeof(string), DataType = "integer")]
        [XmlElement("RowHidden", typeof(ST_TrueFalseBlank))]
        [XmlElement("ScriptExtended", typeof(string))]
        [XmlElement("ScriptLanguage", typeof(string), DataType = "nonNegativeInteger")]
        [XmlElement("ScriptLocation", typeof(string), DataType = "nonNegativeInteger")]
        [XmlElement("ScriptText", typeof(string))]
        [XmlElement("SecretEdit", typeof(ST_TrueFalseBlank))]
        [XmlElement("Sel", typeof(string), DataType = "integer")]
        [XmlElement("SelType", typeof(string))]
        [XmlElement("SizeWithCells", typeof(ST_TrueFalseBlank))]
        [XmlElement("TextHAlign", typeof(string))]
        [XmlElement("TextVAlign", typeof(string))]
        [XmlElement("UIObj", typeof(ST_TrueFalseBlank))]
        [XmlElement("VScroll", typeof(ST_TrueFalseBlank))]
        [XmlElement("VTEdit", typeof(string), DataType = "integer")]
        [XmlElement("Val", typeof(string), DataType = "integer")]
        [XmlElement("ValidIds", typeof(ST_TrueFalseBlank))]
        [XmlElement("Visible", typeof(ST_TrueFalseBlank))]
        [XmlElement("WidthMin", typeof(string), DataType = "integer")]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public object[] Items
        {
            get
            {
                return this.itemsField;
            }
            set
            {
                this.itemsField = value;
            }
        }

        /// <remarks/>
        [XmlElement("ItemsElementName")]
        [XmlIgnore]
        public ItemsChoiceType[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField;
            }
            set
            {
                this.itemsElementNameField = value;
            }
        }
        private List<int> columnField;
        [XmlElement]
        public List<int> column
        {
            get { return this.columnField; }
            set { this.columnField = value; }
        }
        public int GetColumnArray(int index)
        {
            return this.columnField[index];
        }
        private List<int> rowField;

        [XmlElement]
        public List<int> row
        {
            get { return this.rowField; }
            set { this.rowField = value; }
        }

        public int GetRowArray(int index)
        {
            return this.rowField[index];
        }
        /// <remarks/>
        [XmlAttribute]
        public ST_ObjectType ObjectType
        {
            get
            {
                return this.objectTypeField;
            }
            set
            {
                this.objectTypeField = value;
            }
        }
    }

    /// <remarks/>
    [Serializable]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:excel")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:excel", IsNullable = false)]
    public enum ST_TrueFalseBlank
    {

        /// <remarks/>
        @true,

        /// <remarks/>
        t,

        /// <remarks/>
        @false,

        /// <remarks/>
        f,

        /// <remarks/>
        [XmlEnum("")]
        Item,
    }

    /// <remarks/>
    [Serializable]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:excel")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:excel", IsNullable = false)]
    public enum ST_CF
    {

        /// <remarks/>
        PictOld,

        /// <remarks/>
        Pict,

        /// <remarks/>
        Bitmap,

        /// <remarks/>
        PictPrint,

        /// <remarks/>
        PictScreen,
    }

    /// <remarks/>
    [Serializable]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:excel", IncludeInSchema = false)]
    public enum ItemsChoiceType
    {

        /// <remarks/>
        Accel,

        /// <remarks/>
        Accel2,

        /// <remarks/>
        Anchor,

        /// <remarks/>
        AutoFill,

        /// <remarks/>
        AutoLine,

        /// <remarks/>
        AutoPict,

        /// <remarks/>
        AutoScale,

        /// <remarks/>
        CF,

        /// <remarks/>
        Camera,

        /// <remarks/>
        Cancel,

        /// <remarks/>
        Checked,

        /// <remarks/>
        ColHidden,

        /// <remarks/>
        Colored,

        /// <remarks/>
        Column,

        /// <remarks/>
        DDE,

        /// <remarks/>
        Default,

        /// <remarks/>
        DefaultSize,

        /// <remarks/>
        Disabled,

        /// <remarks/>
        Dismiss,

        /// <remarks/>
        DropLines,

        /// <remarks/>
        DropStyle,

        /// <remarks/>
        Dx,

        /// <remarks/>
        FirstButton,

        /// <remarks/>
        FmlaGroup,

        /// <remarks/>
        FmlaLink,

        /// <remarks/>
        FmlaMacro,

        /// <remarks/>
        FmlaPict,

        /// <remarks/>
        FmlaRange,

        /// <remarks/>
        FmlaTxbx,

        /// <remarks/>
        Help,

        /// <remarks/>
        Horiz,

        /// <remarks/>
        Inc,

        /// <remarks/>
        JustLastX,

        /// <remarks/>
        LCT,

        /// <remarks/>
        ListItem,

        /// <remarks/>
        LockText,

        /// <remarks/>
        Locked,

        /// <remarks/>
        MapOCX,

        /// <remarks/>
        Max,

        /// <remarks/>
        Min,

        /// <remarks/>
        MoveWithCells,

        /// <remarks/>
        MultiLine,

        /// <remarks/>
        MultiSel,

        /// <remarks/>
        NoThreeD,

        /// <remarks/>
        NoThreeD2,

        /// <remarks/>
        Page,

        /// <remarks/>
        PrintObject,

        /// <remarks/>
        RecalcAlways,

        /// <remarks/>
        Row,

        /// <remarks/>
        RowHidden,

        /// <remarks/>
        ScriptExtended,

        /// <remarks/>
        ScriptLanguage,

        /// <remarks/>
        ScriptLocation,

        /// <remarks/>
        ScriptText,

        /// <remarks/>
        SecretEdit,

        /// <remarks/>
        Sel,

        /// <remarks/>
        SelType,

        /// <remarks/>
        SizeWithCells,

        /// <remarks/>
        TextHAlign,

        /// <remarks/>
        TextVAlign,

        /// <remarks/>
        UIObj,

        /// <remarks/>
        VScroll,

        /// <remarks/>
        VTEdit,

        /// <remarks/>
        Val,

        /// <remarks/>
        ValidIds,

        /// <remarks/>
        Visible,

        /// <remarks/>
        WidthMin,
    }

    /// <remarks/>
    [Serializable]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:excel")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:excel", IsNullable = false)]
    public enum ST_ObjectType
    {

        /// <remarks/>
        Button,

        /// <remarks/>
        Checkbox,

        /// <remarks/>
        Dialog,

        /// <remarks/>
        Drop,

        /// <remarks/>
        Edit,

        /// <remarks/>
        GBox,

        /// <remarks/>
        Label,

        /// <remarks/>
        LineA,

        /// <remarks/>
        List,

        /// <remarks/>
        Movie,

        /// <remarks/>
        Note,

        /// <remarks/>
        Pict,

        /// <remarks/>
        Radio,

        /// <remarks/>
        RectA,

        /// <remarks/>
        Scroll,

        /// <remarks/>
        Spin,

        /// <remarks/>
        Shape,

        /// <remarks/>
        Group,

        /// <remarks/>
        Rect,
    }
}
