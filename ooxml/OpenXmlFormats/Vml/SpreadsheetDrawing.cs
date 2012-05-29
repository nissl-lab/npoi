using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Vml.Spreadsheet
{

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:excel")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:excel", IsNullable = true)]
    public partial class CT_ClientData
    {

        private List<object> itemsField;

        private ItemsChoiceType[] itemsElementNameField;

        private ST_ObjectType objectTypeField;

    
        //[XmlElement("Accel", typeof(string), DataType = "integer")]
        //[XmlElement("Accel2", typeof(string), DataType = "integer")]
        //[XmlElement("Anchor", typeof(string))]
        //[XmlElement("AutoFill", typeof(ST_TrueFalseBlank))]
        //[XmlElement("AutoLine", typeof(ST_TrueFalseBlank))]
        //[XmlElement("AutoPict", typeof(ST_TrueFalseBlank))]
        //[XmlElement("AutoScale", typeof(ST_TrueFalseBlank))]
        //[XmlElement("CF", typeof(ST_CF))]
        //[XmlElement("Camera", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Cancel", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Checked", typeof(string), DataType = "integer")]
        //[XmlElement("ColHidden", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Colored", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Column", typeof(string), DataType = "integer")]
        //[XmlElement("DDE", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Default", typeof(ST_TrueFalseBlank))]
        //[XmlElement("DefaultSize", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Disabled", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Dismiss", typeof(ST_TrueFalseBlank))]
        //[XmlElement("DropLines", typeof(string), DataType = "integer")]
        //[XmlElement("DropStyle", typeof(string))]
        //[XmlElement("Dx", typeof(string), DataType = "integer")]
        //[XmlElement("FirstButton", typeof(ST_TrueFalseBlank))]
        //[XmlElement("FmlaGroup", typeof(string))]
        //[XmlElement("FmlaLink", typeof(string))]
        //[XmlElement("FmlaMacro", typeof(string))]
        //[XmlElement("FmlaPict", typeof(string))]
        //[XmlElement("FmlaRange", typeof(string))]
        //[XmlElement("FmlaTxbx", typeof(string))]
        //[XmlElement("Help", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Horiz", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Inc", typeof(string), DataType = "integer")]
        //[XmlElement("JustLastX", typeof(ST_TrueFalseBlank))]
        //[XmlElement("LCT", typeof(string))]
        //[XmlElement("ListItem", typeof(string))]
        //[XmlElement("LockText", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Locked", typeof(ST_TrueFalseBlank))]
        //[XmlElement("MapOCX", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Max", typeof(string), DataType = "integer")]
        //[XmlElement("Min", typeof(string), DataType = "integer")]
        //[XmlElement("MoveWithCells", typeof(ST_TrueFalseBlank))]
        //[XmlElement("MultiLine", typeof(ST_TrueFalseBlank))]
        //[XmlElement("MultiSel", typeof(string))]
        //[XmlElement("NoThreeD", typeof(ST_TrueFalseBlank))]
        //[XmlElement("NoThreeD2", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Page", typeof(string), DataType = "integer")]
        //[XmlElement("PrintObject", typeof(ST_TrueFalseBlank))]
        //[XmlElement("RecalcAlways", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Row", typeof(string), DataType = "integer")]
        //[XmlElement("RowHidden", typeof(ST_TrueFalseBlank))]
        //[XmlElement("ScriptExtended", typeof(string))]
        //[XmlElement("ScriptLanguage", typeof(string), DataType = "nonNegativeInteger")]
        //[XmlElement("ScriptLocation", typeof(string), DataType = "nonNegativeInteger")]
        //[XmlElement("ScriptText", typeof(string))]
        //[XmlElement("SecretEdit", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Sel", typeof(string), DataType = "integer")]
        //[XmlElement("SelType", typeof(string))]
        //[XmlElement("SizeWithCells", typeof(ST_TrueFalseBlank))]
        //[XmlElement("TextHAlign", typeof(string))]
        //[XmlElement("TextVAlign", typeof(string))]
        //[XmlElement("UIObj", typeof(ST_TrueFalseBlank))]
        //[XmlElement("VScroll", typeof(ST_TrueFalseBlank))]
        //[XmlElement("VTEdit", typeof(string), DataType = "integer")]
        //[XmlElement("Val", typeof(string), DataType = "integer")]
        //[XmlElement("ValidIds", typeof(ST_TrueFalseBlank))]
        //[XmlElement("Visible", typeof(ST_TrueFalseBlank))]
        //[XmlElement("WidthMin", typeof(string), DataType = "integer")]
        //[XmlChoiceIdentifier("ItemsElementName")]
        //public List<object> Items
        //{
        //    get
        //    {
        //        return this.itemsField;
        //    }
        //    set
        //    {
        //        this.itemsField = value;
        //    }
        //}

    
        //[XmlElement("ItemsElementName")]
        //[XmlIgnore]
        //public ItemsChoiceType[] ItemsElementName
        //{
        //    get
        //    {
        //        return this.itemsElementNameField;
        //    }
        //    set
        //    {
        //        this.itemsElementNameField = value;
        //    }
        //}
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


    [Serializable]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:excel")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:excel", IsNullable = false)]
    public enum ST_TrueFalseBlank
    {
        NONE,    
        @true,

    
        t,

    
        @false,

    
        f,

    
        [XmlEnum("")]
        Item,
    }


    [Serializable]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:excel")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:excel", IsNullable = false)]
    public enum ST_CF
    {

    
        PictOld,

    
        Pict,

    
        Bitmap,

    
        PictPrint,

    
        PictScreen,
    }


    [Serializable]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:excel", IncludeInSchema = false)]
    public enum ItemsChoiceType
    {

    
        Accel,

    
        Accel2,

    
        Anchor,

    
        AutoFill,

    
        AutoLine,

    
        AutoPict,

    
        AutoScale,

    
        CF,

    
        Camera,

    
        Cancel,

    
        Checked,

    
        ColHidden,

    
        Colored,

    
        Column,

    
        DDE,

    
        Default,

    
        DefaultSize,

    
        Disabled,

    
        Dismiss,

    
        DropLines,

    
        DropStyle,

    
        Dx,

    
        FirstButton,

    
        FmlaGroup,

    
        FmlaLink,

    
        FmlaMacro,

    
        FmlaPict,

    
        FmlaRange,

    
        FmlaTxbx,

    
        Help,

    
        Horiz,

    
        Inc,

    
        JustLastX,

    
        LCT,

    
        ListItem,

    
        LockText,

    
        Locked,

    
        MapOCX,

    
        Max,

    
        Min,

    
        MoveWithCells,

    
        MultiLine,

    
        MultiSel,

    
        NoThreeD,

    
        NoThreeD2,

    
        Page,

    
        PrintObject,

    
        RecalcAlways,

    
        Row,

    
        RowHidden,

    
        ScriptExtended,

    
        ScriptLanguage,

    
        ScriptLocation,

    
        ScriptText,

    
        SecretEdit,

    
        Sel,

    
        SelType,

    
        SizeWithCells,

    
        TextHAlign,

    
        TextVAlign,

    
        UIObj,

    
        VScroll,

    
        VTEdit,

    
        Val,

    
        ValidIds,

    
        Visible,

    
        WidthMin,
    }


    [Serializable]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:excel")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:excel", IsNullable = false)]
    public enum ST_ObjectType
    {

    
        Button,

    
        Checkbox,

    
        Dialog,

    
        Drop,

    
        Edit,

    
        GBox,

    
        Label,

    
        LineA,

    
        List,

    
        Movie,

    
        Note,

    
        Pict,

    
        Radio,

    
        RectA,

    
        Scroll,

    
        Spin,

    
        Shape,

    
        Group,

    
        Rect,
    }
}
