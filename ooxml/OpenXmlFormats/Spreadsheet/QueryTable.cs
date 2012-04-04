using System.ComponentModel;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Collections.Generic;
namespace NPOI.OpenXmlFormats.Spreadsheet
{
    
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_QueryTable {
        
        private CT_QueryTableRefresh queryTableRefreshField;
        
        private CT_ExtensionList extLstField;
        
        private string nameField;
        
        private bool headersField;
        
        private bool rowNumbersField;
        
        private bool disableRefreshField;
        
        private bool backgroundRefreshField;
        
        private bool firstBackgroundRefreshField;
        
        private bool refreshOnLoadField;
        
        private ST_GrowShrinkType growShrinkTypeField;
        
        private bool fillFormulasField;
        
        private bool removeDataOnSaveField;
        
        private bool disableEditField;
        
        private bool preserveFormattingField;
        
        private bool adjustColumnWidthField;
        
        private bool intermediateField;
        
        private uint connectionIdField;
        
        private uint autoFormatIdField;
        
        private bool autoFormatIdFieldSpecified;
        
        private bool applyNumberFormatsField;
        
        private bool applyNumberFormatsFieldSpecified;
        
        private bool applyBorderFormatsField;
        
        private bool applyBorderFormatsFieldSpecified;
        
        private bool applyFontFormatsField;
        
        private bool applyFontFormatsFieldSpecified;
        
        private bool applyPatternFormatsField;
        
        private bool applyPatternFormatsFieldSpecified;
        
        private bool applyAlignmentFormatsField;
        
        private bool applyAlignmentFormatsFieldSpecified;
        
        private bool applyWidthHeightFormatsField;
        
        private bool applyWidthHeightFormatsFieldSpecified;
        
        public CT_QueryTable() {
            this.headersField = true;
            this.rowNumbersField = false;
            this.disableRefreshField = false;
            this.backgroundRefreshField = true;
            this.firstBackgroundRefreshField = false;
            this.refreshOnLoadField = false;
            this.growShrinkTypeField = ST_GrowShrinkType.insertDelete;
            this.fillFormulasField = false;
            this.removeDataOnSaveField = false;
            this.disableEditField = false;
            this.preserveFormattingField = true;
            this.adjustColumnWidthField = true;
            this.intermediateField = false;
        }
        
        /// <remarks/>
        public CT_QueryTableRefresh queryTableRefresh {
            get {
                return this.queryTableRefreshField;
            }
            set {
                this.queryTableRefreshField = value;
            }
        }
        
        /// <remarks/>
        public CT_ExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
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
        [DefaultValueAttribute(true)]
        public bool headers {
            get {
                return this.headersField;
            }
            set {
                this.headersField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool rowNumbers {
            get {
                return this.rowNumbersField;
            }
            set {
                this.rowNumbersField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool disableRefresh {
            get {
                return this.disableRefreshField;
            }
            set {
                this.disableRefreshField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool backgroundRefresh {
            get {
                return this.backgroundRefreshField;
            }
            set {
                this.backgroundRefreshField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool firstBackgroundRefresh {
            get {
                return this.firstBackgroundRefreshField;
            }
            set {
                this.firstBackgroundRefreshField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool refreshOnLoad {
            get {
                return this.refreshOnLoadField;
            }
            set {
                this.refreshOnLoadField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(ST_GrowShrinkType.insertDelete)]
        public ST_GrowShrinkType growShrinkType {
            get {
                return this.growShrinkTypeField;
            }
            set {
                this.growShrinkTypeField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool fillFormulas {
            get {
                return this.fillFormulasField;
            }
            set {
                this.fillFormulasField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool removeDataOnSave {
            get {
                return this.removeDataOnSaveField;
            }
            set {
                this.removeDataOnSaveField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool disableEdit {
            get {
                return this.disableEditField;
            }
            set {
                this.disableEditField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool preserveFormatting {
            get {
                return this.preserveFormattingField;
            }
            set {
                this.preserveFormattingField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool adjustColumnWidth {
            get {
                return this.adjustColumnWidthField;
            }
            set {
                this.adjustColumnWidthField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool intermediate {
            get {
                return this.intermediateField;
            }
            set {
                this.intermediateField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public uint connectionId {
            get {
                return this.connectionIdField;
            }
            set {
                this.connectionIdField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public uint autoFormatId {
            get {
                return this.autoFormatIdField;
            }
            set {
                this.autoFormatIdField = value;
            }
        }
        
        /// <remarks/>
        [XmlIgnore]
        public bool autoFormatIdSpecified {
            get {
                return this.autoFormatIdFieldSpecified;
            }
            set {
                this.autoFormatIdFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public bool applyNumberFormats {
            get {
                return this.applyNumberFormatsField;
            }
            set {
                this.applyNumberFormatsField = value;
            }
        }
        
        /// <remarks/>
        [XmlIgnore]
        public bool applyNumberFormatsSpecified {
            get {
                return this.applyNumberFormatsFieldSpecified;
            }
            set {
                this.applyNumberFormatsFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public bool applyBorderFormats {
            get {
                return this.applyBorderFormatsField;
            }
            set {
                this.applyBorderFormatsField = value;
            }
        }
        
        /// <remarks/>
        [XmlIgnore]
        public bool applyBorderFormatsSpecified {
            get {
                return this.applyBorderFormatsFieldSpecified;
            }
            set {
                this.applyBorderFormatsFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public bool applyFontFormats {
            get {
                return this.applyFontFormatsField;
            }
            set {
                this.applyFontFormatsField = value;
            }
        }
        
        /// <remarks/>
        [XmlIgnore]
        public bool applyFontFormatsSpecified {
            get {
                return this.applyFontFormatsFieldSpecified;
            }
            set {
                this.applyFontFormatsFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public bool applyPatternFormats {
            get {
                return this.applyPatternFormatsField;
            }
            set {
                this.applyPatternFormatsField = value;
            }
        }
        
        /// <remarks/>
        [XmlIgnore]
        public bool applyPatternFormatsSpecified {
            get {
                return this.applyPatternFormatsFieldSpecified;
            }
            set {
                this.applyPatternFormatsFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public bool applyAlignmentFormats {
            get {
                return this.applyAlignmentFormatsField;
            }
            set {
                this.applyAlignmentFormatsField = value;
            }
        }
        
        /// <remarks/>
        [XmlIgnore]
        public bool applyAlignmentFormatsSpecified {
            get {
                return this.applyAlignmentFormatsFieldSpecified;
            }
            set {
                this.applyAlignmentFormatsFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public bool applyWidthHeightFormats {
            get {
                return this.applyWidthHeightFormatsField;
            }
            set {
                this.applyWidthHeightFormatsField = value;
            }
        }
        
        /// <remarks/>
        [XmlIgnore]
        public bool applyWidthHeightFormatsSpecified {
            get {
                return this.applyWidthHeightFormatsFieldSpecified;
            }
            set {
                this.applyWidthHeightFormatsFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_QueryTableRefresh {
        
        private CT_QueryTableFields queryTableFieldsField;
        
        private CT_QueryTableDeletedFields queryTableDeletedFieldsField;
        
        private CT_SortState sortStateField;
        
        private CT_ExtensionList extLstField;
        
        private bool preserveSortFilterLayoutField;
        
        private bool fieldIdWrappedField;
        
        private bool headersInLastRefreshField;
        
        private byte minimumVersionField;
        
        private uint nextIdField;
        
        private uint unboundColumnsLeftField;
        
        private uint unboundColumnsRightField;
        
        public CT_QueryTableRefresh() {
            this.preserveSortFilterLayoutField = true;
            this.fieldIdWrappedField = false;
            this.headersInLastRefreshField = true;
            this.minimumVersionField = ((byte)(0));
            this.nextIdField = ((uint)(1));
            this.unboundColumnsLeftField = ((uint)(0));
            this.unboundColumnsRightField = ((uint)(0));
        }
        
        /// <remarks/>
        public CT_QueryTableFields queryTableFields {
            get {
                return this.queryTableFieldsField;
            }
            set {
                this.queryTableFieldsField = value;
            }
        }
        
        /// <remarks/>
        public CT_QueryTableDeletedFields queryTableDeletedFields {
            get {
                return this.queryTableDeletedFieldsField;
            }
            set {
                this.queryTableDeletedFieldsField = value;
            }
        }
        
        /// <remarks/>
        public CT_SortState sortState {
            get {
                return this.sortStateField;
            }
            set {
                this.sortStateField = value;
            }
        }
        
        /// <remarks/>
        public CT_ExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool preserveSortFilterLayout {
            get {
                return this.preserveSortFilterLayoutField;
            }
            set {
                this.preserveSortFilterLayoutField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool fieldIdWrapped {
            get {
                return this.fieldIdWrappedField;
            }
            set {
                this.fieldIdWrappedField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(true)]
        public bool headersInLastRefresh {
            get {
                return this.headersInLastRefreshField;
            }
            set {
                this.headersInLastRefreshField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(typeof(byte), "0")]
        public byte minimumVersion {
            get {
                return this.minimumVersionField;
            }
            set {
                this.minimumVersionField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "1")]
        public uint nextId {
            get {
                return this.nextIdField;
            }
            set {
                this.nextIdField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "0")]
        public uint unboundColumnsLeft {
            get {
                return this.unboundColumnsLeftField;
            }
            set {
                this.unboundColumnsLeftField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "0")]
        public uint unboundColumnsRight {
            get {
                return this.unboundColumnsRightField;
            }
            set {
                this.unboundColumnsRightField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_QueryTableFields {
        
        private List<CT_QueryTableField> queryTableFieldField;
        
        private uint countField;
        
        public CT_QueryTableFields() {
            this.countField = ((uint)(0));
            this.queryTableFieldField = new List<CT_QueryTableField>();
        }
        
        /// <remarks/>
        [XmlElement("queryTableField")]
        public List<CT_QueryTableField> queryTableField
        {
            get {
                return this.queryTableFieldField;
            }
            set {
                this.queryTableFieldField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "0")]
        public uint count {
            get {
                return this.countField;
            }
            set {
                this.countField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_QueryTableField {
        
        private CT_ExtensionList extLstField;
        
        private uint idField;
        
        private string nameField;
        
        private bool dataBoundField;
        
        private bool rowNumbersField;
        
        private bool fillFormulasField;
        
        private bool clippedField;
        
        private uint tableColumnIdField;
        
        public CT_QueryTableField() {
            this.dataBoundField = true;
            this.rowNumbersField = false;
            this.fillFormulasField = false;
            this.clippedField = false;
            this.tableColumnIdField = ((uint)(0));
        }
        
        /// <remarks/>
        public CT_ExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public uint id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
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
        [DefaultValueAttribute(true)]
        public bool dataBound {
            get {
                return this.dataBoundField;
            }
            set {
                this.dataBoundField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool rowNumbers {
            get {
                return this.rowNumbersField;
            }
            set {
                this.rowNumbersField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool fillFormulas {
            get {
                return this.fillFormulasField;
            }
            set {
                this.fillFormulasField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(false)]
        public bool clipped {
            get {
                return this.clippedField;
            }
            set {
                this.clippedField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        [DefaultValueAttribute(typeof(uint), "0")]
        public uint tableColumnId {
            get {
                return this.tableColumnIdField;
            }
            set {
                this.tableColumnIdField = value;
            }
        }
    }

    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_QueryTableDeletedFields {
        
        private CT_DeletedField[] deletedFieldField;
        
        private uint countField;
        
        private bool countFieldSpecified;
        
        /// <remarks/>
        [XmlElement("deletedField")]
        public CT_DeletedField[] deletedField {
            get {
                return this.deletedFieldField;
            }
            set {
                this.deletedFieldField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public uint count {
            get {
                return this.countField;
            }
            set {
                this.countField = value;
            }
        }
        
        /// <remarks/>
        [XmlIgnore]
        public bool countSpecified {
            get {
                return this.countFieldSpecified;
            }
            set {
                this.countFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_DeletedField {
        
        private string nameField;
        
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
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=false)]
    public enum ST_GrowShrinkType {
        
        /// <remarks/>
        insertDelete,
        
        /// <remarks/>
        insertClear,
        
        /// <remarks/>
        overwriteClear,
    }
}
