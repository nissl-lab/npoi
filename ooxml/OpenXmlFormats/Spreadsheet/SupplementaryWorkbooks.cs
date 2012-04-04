using System.Xml.Serialization;
using System.ComponentModel;
using System.Diagnostics;
namespace NPOI.OpenXmlFormats.Spreadsheet
{
    
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalLink {
        
        private object itemField;
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_DdeLink {
        
        private CT_DdeItem[] ddeItemsField;
        
        private string ddeServiceField;
        
        private string ddeTopicField;
        
        /// <remarks/>
        [XmlArrayItemAttribute("ddeItem", IsNullable=false)]
        public CT_DdeItem[] ddeItems {
            get {
                return this.ddeItemsField;
            }
            set {
                this.ddeItemsField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public string ddeService {
            get {
                return this.ddeServiceField;
            }
            set {
                this.ddeServiceField = value;
            }
        }
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
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
        
        /// <remarks/>
        public CT_DdeValues values {
            get {
                return this.valuesField;
            }
            set {
                this.valuesField = value;
            }
        }
        
        /// <remarks/>
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
        
        /// <remarks/>
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
        
        /// <remarks/>
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
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
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
        
        /// <remarks/>
        [XmlElement("value")]
        public CT_DdeValue[] value {
            get {
                return this.valueField;
            }
            set {
                this.valueField = value;
            }
        }
        
        /// <remarks/>
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
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
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
        
        /// <remarks/>
        public string val {
            get {
                return this.valField;
            }
            set {
                this.valField = value;
            }
        }
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=false)]
    public enum ST_DdeValueType {
        
        /// <remarks/>
        nil,
        
        /// <remarks/>
        b,
        
        /// <remarks/>
        n,
        
        /// <remarks/>
        e,
        
        /// <remarks/>
        str,
    }
    

    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalBook {
        
        private CT_ExternalSheetName[] sheetNamesField;
        
        private CT_ExternalDefinedName[] definedNamesField;
        
        private CT_ExternalSheetData[] sheetDataSetField;
        
        private string idField;
        
        /// <remarks/>
        [XmlArrayItemAttribute("sheetName", IsNullable=false)]
        public CT_ExternalSheetName[] sheetNames {
            get {
                return this.sheetNamesField;
            }
            set {
                this.sheetNamesField = value;
            }
        }
        
        /// <remarks/>
        [XmlArrayItemAttribute("definedName", IsNullable=false)]
        public CT_ExternalDefinedName[] definedNames {
            get {
                return this.definedNamesField;
            }
            set {
                this.definedNamesField = value;
            }
        }
        
        /// <remarks/>
        [XmlArrayItemAttribute("sheetData", IsNullable=false)]
        public CT_ExternalSheetData[] sheetDataSet {
            get {
                return this.sheetDataSetField;
            }
            set {
                this.sheetDataSetField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalSheetName {
        
        private string valField;
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalDefinedName {
        
        private string nameField;
        
        private string refersToField;
        
        private uint sheetIdField;
        
        private bool sheetIdFieldSpecified;
        
        /// <remarks/>
        [XmlAttribute]
        public string name {
            get {
                return this.nameField;
            }
            set {
                this.nameField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public string refersTo {
            get {
                return this.refersToField;
            }
            set {
                this.refersToField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public uint sheetId {
            get {
                return this.sheetIdField;
            }
            set {
                this.sheetIdField = value;
            }
        }
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
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
        
        /// <remarks/>
        [XmlElement("row")]
        public CT_ExternalRow[] row {
            get {
                return this.rowField;
            }
            set {
                this.rowField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public uint sheetId {
            get {
                return this.sheetIdField;
            }
            set {
                this.sheetIdField = value;
            }
        }
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalRow {
        
        private CT_ExternalCell[] cellField;
        
        private uint rField;
        
        /// <remarks/>
        [XmlElement("cell")]
        public CT_ExternalCell[] cell {
            get {
                return this.cellField;
            }
            set {
                this.cellField = value;
            }
        }
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
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
        
        /// <remarks/>
        public string v {
            get {
                return this.vField;
            }
            set {
                this.vField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public string r {
            get {
                return this.rField;
            }
            set {
                this.rField = value;
            }
        }
        
        /// <remarks/>
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
        
        /// <remarks/>
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
    
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_OleLink {
        
        private CT_OleItem[] oleItemsField;
        
        private string idField;
        
        private string progIdField;
        
        /// <remarks/>
        [XmlArrayItemAttribute("oleItem", IsNullable=false)]
        public CT_OleItem[] oleItems {
            get {
                return this.oleItemsField;
            }
            set {
                this.oleItemsField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttributeAttribute(Form=System.Xml.Schema.XmlSchemaForm.Qualified, Namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
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
        
        /// <remarks/>
        [XmlAttribute]
        public string name {
            get {
                return this.nameField;
            }
            set {
                this.nameField = value;
            }
        }
        
        /// <remarks/>
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
        
        /// <remarks/>
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
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalSheetNames {
        
        private CT_ExternalSheetName[] sheetNameField;
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalDefinedNames {
        
        private CT_ExternalDefinedName[] definedNameField;
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_ExternalSheetDataSet {
        
        private CT_ExternalSheetData[] sheetDataField;
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_DdeItems {
        
        private CT_DdeItem[] ddeItemField;
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_OleItems {
        
        private CT_OleItem[] oleItemField;
        
        /// <remarks/>
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
