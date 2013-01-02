using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Xml.Serialization;
using System.Collections.Generic;

namespace NPOI.OpenXmlFormats.Spreadsheet
{

    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalLink {
        
        private object itemField;
        
    
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
    }
    

    [Serializable]
    [DebuggerStepThrough]
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
    }
    

    [Serializable]
    [DebuggerStepThrough]
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
    [DebuggerStepThrough]
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
    [DebuggerStepThrough]
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
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalBook {
        
        private CT_ExternalSheetName[] sheetNamesField;
        
        private CT_ExternalDefinedName[] definedNamesField;
        
        private CT_ExternalSheetData[] sheetDataSetField;
        
        private string idField;
        
    
        [XmlArrayItem("sheetName", IsNullable=false)]
        public CT_ExternalSheetName[] sheetNames {
            get {
                return this.sheetNamesField;
            }
            set {
                this.sheetNamesField = value;
            }
        }
        
    
        [XmlArrayItem("definedName", IsNullable=false)]
        public CT_ExternalDefinedName[] definedNames {
            get {
                return this.definedNamesField;
            }
            set {
                this.definedNamesField = value;
            }
        }
        
    
        [XmlArrayItem("sheetData", IsNullable=false)]
        public CT_ExternalSheetData[] sheetDataSet {
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
    }
    

    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
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
    }
    

    [Serializable]
    [DebuggerStepThrough]
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
    }
    

    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalSheetData {
        
        private CT_ExternalRow[] rowField;
        
        private uint sheetIdField;
        
        private bool refreshErrorField;
        
        public CT_ExternalSheetData() {
            this.refreshErrorField = false;
        }
        
    
        [XmlElement("row")]
        public CT_ExternalRow[] row {
            get {
                return this.rowField;
            }
            set {
                this.rowField = value;
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
    }
    

    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalRow {
        
        private CT_ExternalCell[] cellField;
        
        private uint rField;
        
    
        [XmlElement("cell")]
        public CT_ExternalCell[] cell {
            get {
                return this.cellField;
            }
            set {
                this.cellField = value;
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
    }
    

    [Serializable]
    [DebuggerStepThrough]
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
    }
    
    

    [Serializable]
    [DebuggerStepThrough]
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
    }
    

    [Serializable]
    [DebuggerStepThrough]
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
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalSheetNames {
        
        private CT_ExternalSheetName[] sheetNameField;
        
    
        [XmlElement("sheetName")]
        public CT_ExternalSheetName[] sheetName {
            get {
                return this.sheetNameField;
            }
            set {
                this.sheetNameField = value;
            }
        }
    }
    

    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalDefinedNames {
        
        private CT_ExternalDefinedName[] definedNameField;
        
    
        [XmlElement("definedName")]
        public CT_ExternalDefinedName[] definedName {
            get {
                return this.definedNameField;
            }
            set {
                this.definedNameField = value;
            }
        }
    }
    

    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalSheetDataSet {
        
        private CT_ExternalSheetData[] sheetDataField;
        
    
        [XmlElement("sheetData")]
        public CT_ExternalSheetData[] sheetData {
            get {
                return this.sheetDataField;
            }
            set {
                this.sheetDataField = value;
            }
        }
    }
    

    [Serializable]
    [DebuggerStepThrough]
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
    [DebuggerStepThrough]
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
