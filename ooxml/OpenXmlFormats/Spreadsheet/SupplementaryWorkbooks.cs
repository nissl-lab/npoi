using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Xml.Serialization;
using System.Collections.Generic;
using NPOI.OpenXmlFormats.Spreadsheet.Document;
using System.Xml;
using NPOI.OpenXml4Net.Util;
using System.IO;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public enum ExternalLinkItem : int
    {
        none = 0,
        externalBook=1,
        ddeLink,
        extLst,
        oleLink
    }

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", 
        IsNullable=true, ElementName = "externalLink")]
    public partial class CT_ExternalLink {
        
        private object itemField;

        public ExternalLinkItem itemType { get; set; }
        [XmlElement("ddeLink", typeof(CT_DdeLink))]
        [XmlElement("extLst", typeof(CT_ExtensionList))]
        [XmlElement("externalBook", typeof(CT_ExternalBook))]
        [XmlElement("oleLink", typeof(CT_OleLink))]
        public object Item {
            get {
                return this.itemField;
            }
            set {
                this.itemField = value;
            }
        }

        private CT_ExternalBook externalBookField;
        public CT_ExternalBook externalBook
        {
            get { return externalBookField; }
            set { externalBookField = value; }
        }

        private CT_DdeLink ddeLinkField;
        public CT_DdeLink ddlLink
        {
            get { return ddeLinkField; }
            set { ddeLinkField = value; }
        }

        private CT_OleLink oleLinkField;
        public CT_OleLink oleLink
        {
            get { return oleLinkField; }
            set { oleLinkField = value; }
        }

        private CT_ExtensionList extLstField;
        public CT_ExtensionList extLst
        {
            get { return extLstField; }
            set { extLstField = value; }
        }

        public static CT_ExternalLink Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ExternalLink ctObj = new CT_ExternalLink();

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "externalBook")
                {
                    ctObj.externalBookField = CT_ExternalBook.Parse(childNode, namespaceManager);
                    ctObj.itemField = ctObj.externalBookField;
                    ctObj.itemType = ExternalLinkItem.externalBook;
                }
                else if (childNode.LocalName == "ddeLink")
                {
                    ctObj.ddeLinkField = CT_DdeLink.Parse(childNode, namespaceManager);
                    ctObj.itemField = ctObj.ddeLinkField;
                    ctObj.itemType = ExternalLinkItem.ddeLink;
                }
                else if (childNode.LocalName == "oleLink")
                {
                    ctObj.oleLinkField = CT_OleLink.Parse(childNode, namespaceManager);
                    ctObj.itemField = ctObj.oleLinkField;
                    ctObj.itemType = ExternalLinkItem.oleLink;
                }
                else if (childNode.LocalName == "extLst")
                {
                    ctObj.extLstField = CT_ExtensionList.Parse(childNode, namespaceManager);
                    ctObj.itemField = ctObj.extLstField;
                    ctObj.itemType = ExternalLinkItem.extLst;
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write(@"<externalLink xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x14"" xmlns:x14=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"">");
            if (this.externalBookField != null)
                this.externalBookField.Write(sw, "externalBook");
            if (this.ddeLinkField != null)
                this.ddeLinkField.Write(sw, "ddeLink");
            if (this.extLstField != null)
                this.extLstField.Write(sw, "extLst");
            if (this.oleLinkField != null)
                this.oleLinkField.Write(sw, "oleLink");
            sw.Write("</externalLink>");
        }

        public void AddNewExternalBook()
        {
            this.externalBookField = new CT_ExternalBook();
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_DdeLink {
        
        private List<CT_DdeItem> ddeItemsField = null; // 0..1
        
        private string ddeServiceField; // 1..1
        
        private string ddeTopicField; // 1..1


        [XmlArray("ddeItems")]
        [XmlArrayItem("ddeItem")]
        public List<CT_DdeItem> ddeItems {
            get {
                return this.ddeItemsField;
            }
            set {
                this.ddeItemsField = value;
            }
        }
        
    
        [XmlAttribute]
        public string ddeService {
            get {
                return this.ddeServiceField;
            }
            set {
                this.ddeServiceField = value;
            }
        }
        
    
        [XmlAttribute]
        public string ddeTopic {
            get {
                return this.ddeTopicField;
            }
            set {
                this.ddeTopicField = value;
            }
        }

        internal static CT_DdeLink Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            throw new NotImplementedException();
        }

        internal void Write(StreamWriter sw, string p)
        {
            throw new NotImplementedException();
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_DdeItem {
        
        private CT_DdeValues valuesField;
        
        private string nameField;
        
        private bool oleField;
        
        private bool adviseField;
        
        private bool preferPicField;
        
        public CT_DdeItem() {
            this.nameField = "0";
            this.oleField = false;
            this.adviseField = false;
            this.preferPicField = false;
        }
        
    
        public CT_DdeValues values {
            get {
                return this.valuesField;
            }
            set {
                this.valuesField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute("0")]
        public string name {
            get {
                return this.nameField;
            }
            set {
                this.nameField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool ole {
            get {
                return this.oleField;
            }
            set {
                this.oleField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool advise {
            get {
                return this.adviseField;
            }
            set {
                this.adviseField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool preferPic {
            get {
                return this.preferPicField;
            }
            set {
                this.preferPicField = value;
            }
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_DdeValues {
        
        private CT_DdeValue[] valueField;
        
        private uint rowsField;
        
        private uint colsField;
        
        public CT_DdeValues() {
            this.rowsField = ((uint)(1));
            this.colsField = ((uint)(1));
        }
        
    
        [XmlElement("value")]
        public CT_DdeValue[] value {
            get {
                return this.valueField;
            }
            set {
                this.valueField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "1")]
        public uint rows {
            get {
                return this.rowsField;
            }
            set {
                this.rowsField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "1")]
        public uint cols {
            get {
                return this.colsField;
            }
            set {
                this.colsField = value;
            }
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_DdeValue {
        
        private string valField;
        
        private ST_DdeValueType tField;
        
        public CT_DdeValue() {
            this.tField = ST_DdeValueType.n;
        }
        
    
        public string val {
            get {
                return this.valField;
            }
            set {
                this.valField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(ST_DdeValueType.n)]
        public ST_DdeValueType t {
            get {
                return this.tField;
            }
            set {
                this.tField = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=false)]
    public enum ST_DdeValueType {
        
    
        nil,
        
    
        b,
        
    
        n,
        
    
        e,
        
    
        str,
    }
    


    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalBook {
        
        private CT_ExternalSheetNames sheetNamesField;
        
        private CT_ExternalDefinedNames definedNamesField;
        
        private CT_ExternalSheetDataSet sheetDataSetField;
        
        private string idField;
        
    
        [XmlArrayItem("sheetName", IsNullable=false)]
        public CT_ExternalSheetNames sheetNames
        {
            get {
                return this.sheetNamesField;
            }
            set {
                this.sheetNamesField = value;
            }
        }
        
    
        [XmlArrayItem("definedName", IsNullable=false)]
        public CT_ExternalDefinedNames definedNames
        {
            get {
                return this.definedNamesField;
            }
            set {
                this.definedNamesField = value;
            }
        }
        
    
        [XmlArrayItem("sheetData", IsNullable=false)]
        public CT_ExternalSheetDataSet sheetDataSet
        {
            get {
                return this.sheetDataSetField;
            }
            set {
                this.sheetDataSetField = value;
            }
        }
        
        
        [XmlAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }

        internal static CT_ExternalBook Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ExternalBook ctObj = new CT_ExternalBook();
            
            ctObj.idField = XmlHelper.ReadString(node.Attributes["id", namespaceManager.LookupNamespace("r")]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "sheetNames")
                    ctObj.sheetNamesField = CT_ExternalSheetNames.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "definedNames")
                    ctObj.definedNamesField = CT_ExternalDefinedNames.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sheetDataSet")
                    ctObj.sheetDataSetField = CT_ExternalSheetDataSet.Parse(childNode, namespaceManager);
            }

            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "r:id", this.idField);
            sw.Write(">");

            if (this.sheetNamesField != null)
                this.sheetNamesField.Write(sw, "sheetNames");
            if (this.definedNamesField != null)
                this.definedNamesField.Write(sw, "definedNames");

            if (this.sheetDataSetField != null)
                this.sheetDataSetField.Write(sw, "sheetDataSet");

            sw.Write(string.Format("</{0}>", nodeName));
        }
    }
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalSheetName {
        private string valField;
        [XmlAttribute]
        public string val {
            get {
                return this.valField;
            }
            set {
                this.valField = value;
            }
        }

        internal static CT_ExternalSheetName Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_ExternalSheetName name = new CT_ExternalSheetName();
            name.val = XmlHelper.ReadString(node.Attributes["val"]);
            return name;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.valField);
            sw.Write("/>");
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalDefinedName {
        
        private string nameField;
        
        private string refersToField;
        
        private uint sheetIdField;
        
        private bool sheetIdFieldSpecified;
        
    
        [XmlAttribute]
        public string name {
            get {
                return this.nameField;
            }
            set {
                this.nameField = value;
            }
        }
        
    
        [XmlAttribute]
        public string refersTo {
            get {
                return this.refersToField;
            }
            set {
                this.refersToField = value;
            }
        }
        
    
        [XmlAttribute]
        public uint sheetId {
            get {
                return this.sheetIdField;
            }
            set {
                this.sheetIdFieldSpecified = true;
                this.sheetIdField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool sheetIdSpecified {
            get {
                return this.sheetIdFieldSpecified;
            }
            set {
                this.sheetIdFieldSpecified = value;
            }
        }

        public bool IsSetSheetId()
        {
            return this.sheetIdFieldSpecified && this.sheetIdField != 0;
        }

        internal static CT_ExternalDefinedName Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_ExternalDefinedName name = new CT_ExternalDefinedName();
            name.nameField = XmlHelper.ReadString(node.Attributes["name"]);
            name.refersToField = XmlHelper.ReadString(node.Attributes["refersTo"]);
            name.sheetIdFieldSpecified = node.Attributes["sheetId"] != null;
            if (name.sheetIdFieldSpecified)
            {
                name.sheetIdField = XmlHelper.ReadUInt(node.Attributes["sheetId"]);
            }
            return name;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "name", this.nameField);
            XmlHelper.WriteAttribute(sw, "refersTo", this.refersToField);
            if(this.sheetIdFieldSpecified)
                XmlHelper.WriteAttribute(sw, "sheetId", this.sheetIdField);
            sw.Write("/>");
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalSheetData {
        
        public CT_ExternalSheetData ()
        {
            rowField = new List<CT_ExternalRow>();
        }
        private List<CT_ExternalRow> rowField;
        
        private uint sheetIdField;
        
        private bool refreshErrorField;


        [XmlElement("row")]
        public CT_ExternalRow[] row
        {
            get
            {
                return this.rowField.ToArray();
            }
            set
            {
                this.rowField.Clear();
                this.rowField.AddRange(value);
            }
        }
        
    
        [XmlAttribute]
        public uint sheetId {
            get {
                return this.sheetIdField;
            }
            set {
                this.sheetIdField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool refreshError {
            get {
                return this.refreshErrorField;
            }
            set {
                this.refreshErrorField = value;
            }
        }

        internal static CT_ExternalSheetData Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_ExternalSheetData sheetData = new CT_ExternalSheetData();
            sheetData.refreshErrorField = XmlHelper.ReadBool(node.Attributes["refreshError"]);
            sheetData.sheetIdField = XmlHelper.ReadUInt(node.Attributes["sheetId"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "row")
                    sheetData.rowField.Add(CT_ExternalRow.Parse(childNode, namespaceManager));
            }
            return sheetData;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "sheetId", this.sheetIdField,true);
            if(this.refreshError)
                XmlHelper.WriteAttribute(sw, "refreshError", this.refreshErrorField);
            sw.Write(">");
            foreach (CT_ExternalRow ctObj in this.rowField)
            {
                ctObj.Write(sw, "row");
            }

            sw.Write(string.Format("</{0}>", nodeName));
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalRow {
        public CT_ExternalRow()
        {
            cellField = new List<CT_ExternalCell>();
        }
        private List<CT_ExternalCell> cellField;
        
        private uint rField;


        [XmlElement("cell")]
        public CT_ExternalCell[] cell
        {
            get
            {
                return this.cellField.ToArray();
            }
            set
            {
                this.cellField.Clear();
                this.cellField.AddRange(value);
            }
        }
        
    
        [XmlAttribute]
        public uint r {
            get {
                return this.rField;
            }
            set {
                this.rField = value;
            }
        }

        internal static CT_ExternalRow Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_ExternalRow row = new CT_ExternalRow();
            row.r = XmlHelper.ReadUInt(node.Attributes["r"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "cell")
                    row.cellField.Add(CT_ExternalCell.Parse(childNode, namespaceManager));
            }
            return row;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "r", this.rField);
            sw.Write(">");
            foreach (CT_ExternalCell ctObj in this.cellField)
            {
                ctObj.Write(sw, "cell");
            }

            sw.Write(string.Format("</{0}>", nodeName));
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalCell {
        
        private string vField;
        
        private string rField;
        
        private ST_CellType tField;
        
        private uint vmField;
        
        public CT_ExternalCell() {
            this.tField = ST_CellType.n;
            this.vmField = ((uint)(0));
        }
        
    
        public string v {
            get {
                return this.vField;
            }
            set {
                this.vField = value;
            }
        }
        
    
        [XmlAttribute]
        public string r {
            get {
                return this.rField;
            }
            set {
                this.rField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(ST_CellType.n)]
        public ST_CellType t {
            get {
                return this.tField;
            }
            set {
                this.tField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "0")]
        public uint vm {
            get {
                return this.vmField;
            }
            set {
                this.vmField = value;
            }
        }

        internal static CT_ExternalCell Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_ExternalCell ctObj = new CT_ExternalCell();
            ctObj.rField = XmlHelper.ReadString(node.Attributes["r"]);
            if (node.Attributes["t"] != null)
                ctObj.tField = (ST_CellType)Enum.Parse(typeof(ST_CellType), node.Attributes["t"].Value);
            ctObj.vm = XmlHelper.ReadUInt(node.Attributes["vm"]);

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "v")
                    ctObj.v = childNode.InnerText;
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "r", this.rField);
            if (this.t != ST_CellType.n)
                XmlHelper.WriteAttribute(sw, "t", this.tField.ToString());
            XmlHelper.WriteAttribute(sw, "vm", this.vmField);

            if (this.v == null)
            {
                sw.Write("/>");
            }
            else
            {
                sw.Write(">");
                sw.Write(string.Format("<v>{0}</v>", XmlHelper.EncodeXml(this.v)));
                sw.Write(string.Format("</{0}>", nodeName));
            }
        }
    }
    
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_OleLink {
        
        private CT_OleItem[] oleItemsField;
        
        private string idField;
        
        private string progIdField;
        
    
        [XmlArrayItem("oleItem", IsNullable=false)]
        public CT_OleItem[] oleItems {
            get {
                return this.oleItemsField;
            }
            set {
                this.oleItemsField = value;
            }
        }
        
    
        [XmlAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
    
        [XmlAttribute]
        public string progId {
            get {
                return this.progIdField;
            }
            set {
                this.progIdField = value;
            }
        }

        internal static CT_OleLink Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            throw new NotImplementedException();
        }

        internal void Write(StreamWriter sw, string p)
        {
            throw new NotImplementedException();
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_OleItem {
        
        private string nameField;
        
        private bool iconField;
        
        private bool adviseField;
        
        private bool preferPicField;
        
        public CT_OleItem() {
            this.iconField = false;
            this.adviseField = false;
            this.preferPicField = false;
        }
        
    
        [XmlAttribute]
        public string name {
            get {
                return this.nameField;
            }
            set {
                this.nameField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool icon {
            get {
                return this.iconField;
            }
            set {
                this.iconField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool advise {
            get {
                return this.adviseField;
            }
            set {
                this.adviseField = value;
            }
        }
        
    
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool preferPic {
            get {
                return this.preferPicField;
            }
            set {
                this.preferPicField = value;
            }
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalSheetNames {
        
        public CT_ExternalSheetNames()
        {
            this.sheetNameField = new List<CT_ExternalSheetName>();
        }

        private List<CT_ExternalSheetName> sheetNameField;


        [XmlElement("sheetName")]
        public CT_ExternalSheetName[] sheetName
        {
            get
            {
                return this.sheetNameField.ToArray();
            }
            set
            {
                this.sheetNameField.Clear();
                this.sheetNameField.AddRange(value);
            }
        }

        internal static CT_ExternalSheetNames Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_ExternalSheetNames ctObj = new CT_ExternalSheetNames();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                ctObj.sheetNameField.Add(CT_ExternalSheetName.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");

            foreach (CT_ExternalSheetName ctObj in this.sheetNameField)
            {
                ctObj.Write(sw, "sheetName");
            }

            sw.Write(string.Format("</{0}>", nodeName));
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalDefinedNames {

        public CT_ExternalDefinedNames()
        {
            definedNameField = new List<CT_ExternalDefinedName>();
        }
        private List<CT_ExternalDefinedName> definedNameField;


        [XmlElement("definedName")]
        public CT_ExternalDefinedName[] definedName
        {
            get
            {
                return this.definedNameField.ToArray();
            }
            set
            {
                this.definedNameField.Clear();
                this.definedNameField.AddRange(value);
            }
        }

        internal static CT_ExternalDefinedNames Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_ExternalDefinedNames ctObj = new CT_ExternalDefinedNames();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                ctObj.definedNameField.Add(CT_ExternalDefinedName.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");

            foreach (CT_ExternalDefinedName ctObj in this.definedNameField)
            {
                ctObj.Write(sw, "definedName");
            }

            sw.Write(string.Format("</{0}>", nodeName));
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalSheetDataSet {
        public CT_ExternalSheetDataSet()
        {
            sheetDataField = new List<CT_ExternalSheetData>();
        }
        private List<CT_ExternalSheetData> sheetDataField;


        [XmlElement("sheetData")]
        public CT_ExternalSheetData[] sheetData
        {
            get
            {
                return this.sheetDataField.ToArray();
            }
            set
            {
                this.sheetDataField.Clear();
                this.sheetDataField.AddRange(value);
            }
        }

        internal static CT_ExternalSheetDataSet Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_ExternalSheetDataSet ctObj = new CT_ExternalSheetDataSet();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                ctObj.sheetDataField.Add(CT_ExternalSheetData.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            foreach (CT_ExternalSheetData ctObj in this.sheetDataField)
            {
                ctObj.Write(sw, "sheetData");
            }

            sw.Write(string.Format("</{0}>", nodeName));
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_DdeItems {
        
        private CT_DdeItem[] ddeItemField;
        
    
        [XmlElement("ddeItem")]
        public CT_DdeItem[] ddeItem {
            get {
                return this.ddeItemField;
            }
            set {
                this.ddeItemField = value;
            }
        }
    }
    

    [Serializable]
    
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_OleItems {
        
        private CT_OleItem[] oleItemField;
        
    
        [XmlElement("oleItem")]
        public CT_OleItem[] oleItem {
            get {
                return this.oleItemField;
            }
            set {
                this.oleItemField = value;
            }
        }
    }
}
