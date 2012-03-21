namespace NPOI.OpenXmlFormats.Spreadsheet {
    
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute("MapInfo",Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public class CT_MapInfo {
        
        private CT_Schema[] schemaField;
        
        private CT_Map[] mapField;
        
        private string selectionNamespacesField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Schema")]
        public CT_Schema[] Schema {
            get {
                return this.schemaField;
            }
            set {
                this.schemaField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Map")]
        public CT_Map[] Map {
            get {
                return this.mapField;
            }
            set {
                this.mapField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string SelectionNamespaces {
            get {
                return this.selectionNamespacesField;
            }
            set {
                this.selectionNamespacesField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_Schema {
        
        private System.Xml.XmlElement anyField;
        
        private string idField;
        
        private string schemaRefField;
        
        private string namespaceField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAnyElementAttribute()]
        public System.Xml.XmlElement Any {
            get {
                return this.anyField;
            }
            set {
                this.anyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ID {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string SchemaRef {
            get {
                return this.schemaRefField;
            }
            set {
                this.schemaRefField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Namespace {
            get {
                return this.namespaceField;
            }
            set {
                this.namespaceField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_Map {
        
        private CT_DataBinding dataBindingField;
        
        private uint idField;
        
        private string nameField;
        
        private string rootElementField;
        
        private string schemaIDField;
        
        private bool showImportExportValidationErrorsField;
        
        private bool autoFitField;
        
        private bool appendField;
        
        private bool preserveSortAFLayoutField;
        
        private bool preserveFormatField;
        
        /// <remarks/>
        public CT_DataBinding DataBinding {
            get {
                return this.dataBindingField;
            }
            set {
                this.dataBindingField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint ID {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string Name {
            get {
                return this.nameField;
            }
            set {
                this.nameField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string RootElement {
            get {
                return this.rootElementField;
            }
            set {
                this.rootElementField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string SchemaID {
            get {
                return this.schemaIDField;
            }
            set {
                this.schemaIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool ShowImportExportValidationErrors {
            get {
                return this.showImportExportValidationErrorsField;
            }
            set {
                this.showImportExportValidationErrorsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool AutoFit {
            get {
                return this.autoFitField;
            }
            set {
                this.autoFitField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool Append {
            get {
                return this.appendField;
            }
            set {
                this.appendField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool PreserveSortAFLayout {
            get {
                return this.preserveSortAFLayoutField;
            }
            set {
                this.preserveSortAFLayoutField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool PreserveFormat {
            get {
                return this.preserveFormatField;
            }
            set {
                this.preserveFormatField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_DataBinding {
        
        private System.Xml.XmlElement anyField;
        
        private string dataBindingNameField;
        
        private bool fileBindingField;
        
        private bool fileBindingFieldSpecified;
        
        private uint connectionIDField;
        
        private bool connectionIDFieldSpecified;
        
        private string fileBindingNameField;
        
        private uint dataBindingLoadModeField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAnyElementAttribute()]
        public System.Xml.XmlElement Any {
            get {
                return this.anyField;
            }
            set {
                this.anyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string DataBindingName {
            get {
                return this.dataBindingNameField;
            }
            set {
                this.dataBindingNameField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool FileBinding {
            get {
                return this.fileBindingField;
            }
            set {
                this.fileBindingField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool FileBindingSpecified {
            get {
                return this.fileBindingFieldSpecified;
            }
            set {
                this.fileBindingFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint ConnectionID {
            get {
                return this.connectionIDField;
            }
            set {
                this.connectionIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool ConnectionIDSpecified {
            get {
                return this.connectionIDFieldSpecified;
            }
            set {
                this.connectionIDFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string FileBindingName {
            get {
                return this.fileBindingNameField;
            }
            set {
                this.fileBindingNameField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint DataBindingLoadMode {
            get {
                return this.dataBindingLoadModeField;
            }
            set {
                this.dataBindingLoadModeField = value;
            }
        }
    }
}
