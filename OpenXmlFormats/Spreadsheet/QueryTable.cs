using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    
    

    [Serializable]
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
        
    
        public CT_QueryTableRefresh queryTableRefresh {
            get {
                return this.queryTableRefreshField;
            }
            set {
                this.queryTableRefreshField = value;
            }
        }
        
    
        public CT_ExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
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
        [DefaultValueAttribute(true)]
        public bool headers {
            get {
                return this.headersField;
            }
            set {
                this.headersField = value;
            }
        }
        
    
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
        
    
        [XmlAttribute]
        public uint connectionId {
            get {
                return this.connectionIdField;
            }
            set {
                this.connectionIdField = value;
            }
        }
        
    
        [XmlAttribute]
        public uint autoFormatId {
            get {
                return this.autoFormatIdField;
            }
            set {
                this.autoFormatIdField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool autoFormatIdSpecified {
            get {
                return this.autoFormatIdFieldSpecified;
            }
            set {
                this.autoFormatIdFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool applyNumberFormats {
            get {
                return this.applyNumberFormatsField;
            }
            set {
                this.applyNumberFormatsField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool applyNumberFormatsSpecified {
            get {
                return this.applyNumberFormatsFieldSpecified;
            }
            set {
                this.applyNumberFormatsFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool applyBorderFormats {
            get {
                return this.applyBorderFormatsField;
            }
            set {
                this.applyBorderFormatsField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool applyBorderFormatsSpecified {
            get {
                return this.applyBorderFormatsFieldSpecified;
            }
            set {
                this.applyBorderFormatsFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool applyFontFormats {
            get {
                return this.applyFontFormatsField;
            }
            set {
                this.applyFontFormatsField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool applyFontFormatsSpecified {
            get {
                return this.applyFontFormatsFieldSpecified;
            }
            set {
                this.applyFontFormatsFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool applyPatternFormats {
            get {
                return this.applyPatternFormatsField;
            }
            set {
                this.applyPatternFormatsField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool applyPatternFormatsSpecified {
            get {
                return this.applyPatternFormatsFieldSpecified;
            }
            set {
                this.applyPatternFormatsFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool applyAlignmentFormats {
            get {
                return this.applyAlignmentFormatsField;
            }
            set {
                this.applyAlignmentFormatsField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool applyAlignmentFormatsSpecified {
            get {
                return this.applyAlignmentFormatsFieldSpecified;
            }
            set {
                this.applyAlignmentFormatsFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public bool applyWidthHeightFormats {
            get {
                return this.applyWidthHeightFormatsField;
            }
            set {
                this.applyWidthHeightFormatsField = value;
            }
        }
        
    
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
    

    [Serializable]
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
        
    
        public CT_QueryTableFields queryTableFields {
            get {
                return this.queryTableFieldsField;
            }
            set {
                this.queryTableFieldsField = value;
            }
        }
        
    
        public CT_QueryTableDeletedFields queryTableDeletedFields {
            get {
                return this.queryTableDeletedFieldsField;
            }
            set {
                this.queryTableDeletedFieldsField = value;
            }
        }
        
    
        public CT_SortState sortState {
            get {
                return this.sortStateField;
            }
            set {
                this.sortStateField = value;
            }
        }
        
    
        public CT_ExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
    
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
    

    [Serializable]
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
    

    [Serializable]
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
        
    
        public CT_ExtensionList extLst {
            get {
                return this.extLstField;
            }
            set {
                this.extLstField = value;
            }
        }
        
    
        [XmlAttribute]
        public uint id {
            get {
                return this.idField;
            }
            set {
                this.idField = value;
            }
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
        [DefaultValueAttribute(true)]
        public bool dataBound {
            get {
                return this.dataBoundField;
            }
            set {
                this.dataBoundField = value;
            }
        }
        
    
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


    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_QueryTableDeletedFields {
        
        private CT_DeletedField[] deletedFieldField;
        
        private uint countField;
        
        private bool countFieldSpecified;
        
    
        [XmlElement("deletedField")]
        public CT_DeletedField[] deletedField {
            get {
                return this.deletedFieldField;
            }
            set {
                this.deletedFieldField = value;
            }
        }
        
    
        [XmlAttribute]
        public uint count {
            get {
                return this.countField;
            }
            set {
                this.countField = value;
            }
        }
        
    
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
    

    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=true)]
    public partial class CT_DeletedField {
        
        private string nameField;
        
    
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
    

    [Serializable]
    [XmlType(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable=false)]
    public enum ST_GrowShrinkType {
        
    
        insertDelete,
        
    
        insertClear,
        
    
        overwriteClear,
    }
}
