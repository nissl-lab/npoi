using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.ComponentModel;
using System.IO;
using System.Xml;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Vml.Spreadsheet
{

    [Serializable]

    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:excel")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:excel", IsNullable = true)]
    public class CT_ClientData
    {
        public CT_ClientData()
        {
            this.rowField = new List<int>();
            this.columnField = new List<int>();
        }
        //private List<object> itemsField;

        //private ItemsChoiceType[] itemsElementNameField;

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
        public static CT_ClientData Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ClientData ctObj = new CT_ClientData();
            if (node.Attributes["ObjectType"] != null)
                ctObj.ObjectType = (ST_ObjectType)Enum.Parse(typeof(ST_ObjectType), node.Attributes["ObjectType"].Value);
            ctObj.column = new List<Int32>();
            ctObj.row = new List<Int32>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "Anchor")
                    ctObj.anchor = childNode.InnerText;
                else if (childNode.LocalName == "AutoFill")
                    ctObj.autoFill = NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalseBlank(childNode.InnerText);
                else if (childNode.LocalName == "Visible")
                    ctObj.visible =NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalseBlank(childNode.InnerText);
                else if (childNode.LocalName == "MoveWithCells")
                    ctObj.moveWithCells =NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalseBlank(childNode.InnerText);
                else if (childNode.LocalName == "SizeWithCells")
                    ctObj.sizeWithCells =NPOI.OpenXmlFormats.Util.XmlHelper.ReadTrueFalseBlank(childNode.InnerText);
                else if (childNode.LocalName == "Column")
                    ctObj.column.Add(Int32.Parse(childNode.InnerText));
                else if (childNode.LocalName == "Row")
                    ctObj.row.Add(Int32.Parse(childNode.InnerText));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<x:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "ObjectType", this.ObjectType.ToString());
            sw.Write(">");
            if (this.moveWithCells == ST_TrueFalseBlank.t || this.moveWithCells == ST_TrueFalseBlank.@true)
                sw.Write(string.Format("<x:MoveWithCells/>", this.moveWithCells));
            if (this.sizeWithCells == ST_TrueFalseBlank.t || this.sizeWithCells == ST_TrueFalseBlank.@true)
                sw.Write(string.Format("<x:SizeWithCells/>", this.sizeWithCells));
            if (this.anchor != null)
                sw.Write(string.Format("<x:Anchor>{0}</x:Anchor>", this.anchor));
            if (this.autoFill != ST_TrueFalseBlank.NONE)
                sw.Write(string.Format("<x:AutoFill>{0}</x:AutoFill>", this.autoFill));
            if (this.visible != ST_TrueFalseBlank.NONE)
                sw.Write(string.Format("<x:Visible>{0}</x:Visible>", this.visible));
            if (this.row != null)
            {
                foreach (Int32 x in this.row)
                {
                    sw.Write(string.Format("<x:Row>{0}</x:Row>", x));
                }
            }
            if (this.column != null)
            {
                foreach (Int32 x in this.column)
                {
                    sw.Write(string.Format("<x:Column>{0}</x:Column>", x));
                }
            }
            sw.Write(string.Format("</x:{0}>", nodeName));
        }


        public void AddNewRow(int rowNum)
        {
            if (rowField != null)
            {
                rowField.Add(rowNum);
            }
        }
        public void AddNewColumn(int columnNum)
        {
            if (columnField != null)
            {
                columnField.Add(columnNum);
            }        
        }

        public void AddNewMoveWithCells()
        {
            this.moveWithCellsField = ST_TrueFalseBlank.t;
            this.moveWithCellsFieldSpecified = true;
        }
        public void AddNewSizeWithCells()
        {
            this.sizeWithCellsField = ST_TrueFalseBlank.t;
            this.sizeWithCellsFieldSpecified = true;
        }
        private string anchorField;
        [XmlElement(ElementName = "Anchor")]
        public string anchor
        {
            get { return this.anchorField; }
            set { this.anchorField = value; }
        }

        public void AddNewAnchor(string name)
        {
            this.anchorField = name;
        }

        public void AddNewAutoFill(ST_TrueFalseBlank value)
        {
            this.autoFillField = value;
            this.autoFillFieldSpecified = true;
        }

        ST_TrueFalseBlank autoFillField = ST_TrueFalseBlank.NONE;
        bool autoFillFieldSpecified = false;

        [XmlElement(ElementName = "AutoFill")]
        [DefaultValue(ST_TrueFalseBlank.NONE)]
        public ST_TrueFalseBlank autoFill
        {
            get { return this.autoFillField; }
            set { this.autoFillField = value; }
        }
        [XmlIgnore]
        public bool autoFillSpecified
        {
            get { return this.autoFillFieldSpecified; }
            set { this.autoFillFieldSpecified = value; }
        }

        ST_TrueFalseBlank visibleField = ST_TrueFalseBlank.NONE;
        bool visibleFieldSpecified = false;

        [XmlElement(ElementName = "Visible")]
        [DefaultValue(ST_TrueFalseBlank.NONE)]
        public ST_TrueFalseBlank visible
        {
            get { return this.visibleField; }
            set { this.visibleField = value; }
        }
        [XmlIgnore]
        public bool visibleSpecified
        {
            get { return this.visibleFieldSpecified; }
            set { this.visibleFieldSpecified = value; }
        }

        ST_TrueFalseBlank moveWithCellsField= ST_TrueFalseBlank.NONE;
        bool moveWithCellsFieldSpecified = false;

        [XmlElement(ElementName="MoveWithCells")]
        [DefaultValue(ST_TrueFalseBlank.NONE)]
        public ST_TrueFalseBlank moveWithCells
        {
            get { return this.moveWithCellsField; }
            set { this.moveWithCellsField = value; }
        }
        [XmlIgnore]
        public bool moveWithCellsSpecified
        {
            get { return this.moveWithCellsFieldSpecified; }
            set { this.moveWithCellsFieldSpecified = value; }
        }
        public int SizeOfMoveWithCellsArray()
        {
            return moveWithCellsSpecified?1:0;
        }
        public int SizeOfSizeWithCellsArray()
        {
            return sizeWithCellsFieldSpecified?1:0;
        }
        ST_TrueFalseBlank sizeWithCellsField = ST_TrueFalseBlank.NONE;
        bool sizeWithCellsFieldSpecified = false;

        [XmlElement(ElementName = "SizeWithCells")]
        [DefaultValue(ST_TrueFalseBlank.NONE)]
        public ST_TrueFalseBlank sizeWithCells
        {
            get { return this.sizeWithCellsField; }
            set { this.sizeWithCellsField = value; }
        }
        [XmlIgnore]
        public bool sizeWithCellsSpecified
        {
            get { return this.sizeWithCellsFieldSpecified; }
            set { this.sizeWithCellsFieldSpecified = value; }
        }
        

        private List<int> columnField;
        [XmlElement(ElementName = "Column")]
        public List<int> column
        {
            get { return this.columnField; }
            set { this.columnField = value; }
        }
        public int GetColumnArray(int index)
        {
            return this.columnField[index];
        }
        public void SetColumnArray(int index, int value)
        {
            this.columnField[index]= value;
        }
        public void SetRowArray(int index, int value)
        {
            this.rowField[index] = value;
        }
        public void SetAnchorArray(int index, string value)
        {
            AddNewAnchor(value);
        }
        private List<int> rowField;

        [XmlElement(ElementName="Row")]
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

        public string GetAnchorArray(int p)
        {
            return this.anchor;
        }
    }


    [Serializable]
    [XmlType(Namespace = "urn:schemas-microsoft-com:office:excel")]
    [XmlRoot(Namespace = "urn:schemas-microsoft-com:office:excel", IsNullable = false)]
    public enum ST_TrueFalseBlank
    {
        NONE,
        [XmlEnum("True")]
        @true,

    
        t,

        [XmlEnum("False")]
        @false,

    
        f,

    
        //[XmlEnum("")]
        //Item,
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
