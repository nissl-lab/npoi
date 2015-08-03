using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("pivotTableDefinition", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public partial class CT_PivotTableDefinition
    {

        private CT_Location locationField;

        private CT_PivotFields pivotFieldsField;

        private CT_RowFields rowFieldsField;

        private CT_rowItems rowItemsField;

        private CT_ColFields colFieldsField;

        private CT_colItems colItemsField;

        private CT_PageFields pageFieldsField;

        private CT_DataFields dataFieldsField;

        private CT_Formats formatsField;

        private CT_ConditionalFormats conditionalFormatsField;

        private CT_ChartFormats chartFormatsField;

        private CT_PivotHierarchies pivotHierarchiesField;

        private CT_PivotTableStyle pivotTableStyleInfoField;

        private CT_PivotFilters filtersField;

        private CT_RowHierarchiesUsage rowHierarchiesUsageField;

        private CT_ColHierarchiesUsage colHierarchiesUsageField;

        private CT_ExtensionList extLstField;

        private string nameField;

        private uint cacheIdField;

        private bool dataOnRowsField;

        private uint dataPositionField;

        private bool dataPositionFieldSpecified;

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

        private string dataCaptionField;

        private string grandTotalCaptionField;

        private string errorCaptionField;

        private bool showErrorField;

        private string missingCaptionField;

        private bool showMissingField;

        private string pageStyleField;

        private string pivotTableStyleField;

        private string vacatedStyleField;

        private string tagField;

        private byte updatedVersionField;

        private byte minRefreshableVersionField;

        private bool asteriskTotalsField;

        private bool showItemsField;

        private bool editDataField;

        private bool disableFieldListField;

        private bool showCalcMbrsField;

        private bool visualTotalsField;

        private bool showMultipleLabelField;

        private bool showDataDropDownField;

        private bool showDrillField;

        private bool printDrillField;

        private bool showMemberPropertyTipsField;

        private bool showDataTipsField;

        private bool enableWizardField;

        private bool enableDrillField;

        private bool enableFieldPropertiesField;

        private bool preserveFormattingField;

        private bool useAutoFormattingField;

        private uint pageWrapField;

        private bool pageOverThenDownField;

        private bool subtotalHiddenItemsField;

        private bool rowGrandTotalsField;

        private bool colGrandTotalsField;

        private bool fieldPrintTitlesField;

        private bool itemPrintTitlesField;

        private bool mergeItemField;

        private bool showDropZonesField;

        private byte createdVersionField;

        private uint indentField;

        private bool showEmptyRowField;

        private bool showEmptyColField;

        private bool showHeadersField;

        private bool compactField;

        private bool outlineField;

        private bool outlineDataField;

        private bool compactDataField;

        private bool publishedField;

        private bool gridDropZonesField;

        private bool immersiveField;

        private bool multipleFieldFiltersField;

        private uint chartFormatField;

        private string rowHeaderCaptionField;

        private string colHeaderCaptionField;

        private bool fieldListSortAscendingField;

        private bool mdxSubqueriesField;

        private bool customListSortField;

        public CT_PivotTableDefinition()
        {
            this.extLstField = new CT_ExtensionList();
            this.colHierarchiesUsageField = new CT_ColHierarchiesUsage();
            this.rowHierarchiesUsageField = new CT_RowHierarchiesUsage();
            this.filtersField = new CT_PivotFilters();
            this.pivotTableStyleInfoField = new CT_PivotTableStyle();
            this.pivotHierarchiesField = new CT_PivotHierarchies();
            this.chartFormatsField = new CT_ChartFormats();
            this.conditionalFormatsField = new CT_ConditionalFormats();
            this.formatsField = new CT_Formats();
            this.dataFieldsField = new CT_DataFields();
            this.pageFieldsField = new CT_PageFields();
            this.colItemsField = new CT_colItems();
            this.colFieldsField = new CT_ColFields();
            this.rowItemsField = new CT_rowItems();
            this.rowFieldsField = new CT_RowFields();
            this.pivotFieldsField = new CT_PivotFields();
            this.locationField = new CT_Location();
            this.dataOnRowsField = false;
            this.showErrorField = false;
            this.showMissingField = true;
            this.updatedVersionField = ((byte)(0));
            this.minRefreshableVersionField = ((byte)(0));
            this.asteriskTotalsField = false;
            this.showItemsField = true;
            this.editDataField = false;
            this.disableFieldListField = false;
            this.showCalcMbrsField = true;
            this.visualTotalsField = true;
            this.showMultipleLabelField = true;
            this.showDataDropDownField = true;
            this.showDrillField = true;
            this.printDrillField = false;
            this.showMemberPropertyTipsField = true;
            this.showDataTipsField = true;
            this.enableWizardField = true;
            this.enableDrillField = true;
            this.enableFieldPropertiesField = true;
            this.preserveFormattingField = true;
            this.useAutoFormattingField = false;
            this.pageWrapField = ((uint)(0));
            this.pageOverThenDownField = false;
            this.subtotalHiddenItemsField = false;
            this.rowGrandTotalsField = true;
            this.colGrandTotalsField = true;
            this.fieldPrintTitlesField = false;
            this.itemPrintTitlesField = false;
            this.mergeItemField = false;
            this.showDropZonesField = true;
            this.createdVersionField = ((byte)(0));
            this.indentField = ((uint)(1));
            this.showEmptyRowField = false;
            this.showEmptyColField = false;
            this.showHeadersField = true;
            this.compactField = true;
            this.outlineField = false;
            this.outlineDataField = false;
            this.compactDataField = true;
            this.publishedField = false;
            this.gridDropZonesField = false;
            this.immersiveField = true;
            this.multipleFieldFiltersField = true;
            this.chartFormatField = ((uint)(0));
            this.fieldListSortAscendingField = false;
            this.mdxSubqueriesField = false;
            this.customListSortField = true;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_Location location
        {
            get
            {
                return this.locationField;
            }
            set
            {
                this.locationField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_PivotFields pivotFields
        {
            get
            {
                return this.pivotFieldsField;
            }
            set
            {
                this.pivotFieldsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
        public CT_RowFields rowFields
        {
            get
            {
                return this.rowFieldsField;
            }
            set
            {
                this.rowFieldsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 3)]
        public CT_rowItems rowItems
        {
            get
            {
                return this.rowItemsField;
            }
            set
            {
                this.rowItemsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 4)]
        public CT_ColFields colFields
        {
            get
            {
                return this.colFieldsField;
            }
            set
            {
                this.colFieldsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 5)]
        public CT_colItems colItems
        {
            get
            {
                return this.colItemsField;
            }
            set
            {
                this.colItemsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 6)]
        public CT_PageFields pageFields
        {
            get
            {
                return this.pageFieldsField;
            }
            set
            {
                this.pageFieldsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 7)]
        public CT_DataFields dataFields
        {
            get
            {
                return this.dataFieldsField;
            }
            set
            {
                this.dataFieldsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 8)]
        public CT_Formats formats
        {
            get
            {
                return this.formatsField;
            }
            set
            {
                this.formatsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 9)]
        public CT_ConditionalFormats conditionalFormats
        {
            get
            {
                return this.conditionalFormatsField;
            }
            set
            {
                this.conditionalFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 10)]
        public CT_ChartFormats chartFormats
        {
            get
            {
                return this.chartFormatsField;
            }
            set
            {
                this.chartFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 11)]
        public CT_PivotHierarchies pivotHierarchies
        {
            get
            {
                return this.pivotHierarchiesField;
            }
            set
            {
                this.pivotHierarchiesField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 12)]
        public CT_PivotTableStyle pivotTableStyleInfo
        {
            get
            {
                return this.pivotTableStyleInfoField;
            }
            set
            {
                this.pivotTableStyleInfoField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 13)]
        public CT_PivotFilters filters
        {
            get
            {
                return this.filtersField;
            }
            set
            {
                this.filtersField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 14)]
        public CT_RowHierarchiesUsage rowHierarchiesUsage
        {
            get
            {
                return this.rowHierarchiesUsageField;
            }
            set
            {
                this.rowHierarchiesUsageField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 15)]
        public CT_ColHierarchiesUsage colHierarchiesUsage
        {
            get
            {
                return this.colHierarchiesUsageField;
            }
            set
            {
                this.colHierarchiesUsageField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 16)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint cacheId
        {
            get
            {
                return this.cacheIdField;
            }
            set
            {
                this.cacheIdField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool dataOnRows
        {
            get
            {
                return this.dataOnRowsField;
            }
            set
            {
                this.dataOnRowsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint dataPosition
        {
            get
            {
                return this.dataPositionField;
            }
            set
            {
                this.dataPositionField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool dataPositionSpecified
        {
            get
            {
                return this.dataPositionFieldSpecified;
            }
            set
            {
                this.dataPositionFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint autoFormatId
        {
            get
            {
                return this.autoFormatIdField;
            }
            set
            {
                this.autoFormatIdField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool autoFormatIdSpecified
        {
            get
            {
                return this.autoFormatIdFieldSpecified;
            }
            set
            {
                this.autoFormatIdFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyNumberFormats
        {
            get
            {
                return this.applyNumberFormatsField;
            }
            set
            {
                this.applyNumberFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyNumberFormatsSpecified
        {
            get
            {
                return this.applyNumberFormatsFieldSpecified;
            }
            set
            {
                this.applyNumberFormatsFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyBorderFormats
        {
            get
            {
                return this.applyBorderFormatsField;
            }
            set
            {
                this.applyBorderFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyBorderFormatsSpecified
        {
            get
            {
                return this.applyBorderFormatsFieldSpecified;
            }
            set
            {
                this.applyBorderFormatsFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyFontFormats
        {
            get
            {
                return this.applyFontFormatsField;
            }
            set
            {
                this.applyFontFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyFontFormatsSpecified
        {
            get
            {
                return this.applyFontFormatsFieldSpecified;
            }
            set
            {
                this.applyFontFormatsFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyPatternFormats
        {
            get
            {
                return this.applyPatternFormatsField;
            }
            set
            {
                this.applyPatternFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyPatternFormatsSpecified
        {
            get
            {
                return this.applyPatternFormatsFieldSpecified;
            }
            set
            {
                this.applyPatternFormatsFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyAlignmentFormats
        {
            get
            {
                return this.applyAlignmentFormatsField;
            }
            set
            {
                this.applyAlignmentFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyAlignmentFormatsSpecified
        {
            get
            {
                return this.applyAlignmentFormatsFieldSpecified;
            }
            set
            {
                this.applyAlignmentFormatsFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyWidthHeightFormats
        {
            get
            {
                return this.applyWidthHeightFormatsField;
            }
            set
            {
                this.applyWidthHeightFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyWidthHeightFormatsSpecified
        {
            get
            {
                return this.applyWidthHeightFormatsFieldSpecified;
            }
            set
            {
                this.applyWidthHeightFormatsFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string dataCaption
        {
            get
            {
                return this.dataCaptionField;
            }
            set
            {
                this.dataCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string grandTotalCaption
        {
            get
            {
                return this.grandTotalCaptionField;
            }
            set
            {
                this.grandTotalCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string errorCaption
        {
            get
            {
                return this.errorCaptionField;
            }
            set
            {
                this.errorCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showError
        {
            get
            {
                return this.showErrorField;
            }
            set
            {
                this.showErrorField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string missingCaption
        {
            get
            {
                return this.missingCaptionField;
            }
            set
            {
                this.missingCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showMissing
        {
            get
            {
                return this.showMissingField;
            }
            set
            {
                this.showMissingField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string pageStyle
        {
            get
            {
                return this.pageStyleField;
            }
            set
            {
                this.pageStyleField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string pivotTableStyle
        {
            get
            {
                return this.pivotTableStyleField;
            }
            set
            {
                this.pivotTableStyleField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string vacatedStyle
        {
            get
            {
                return this.vacatedStyleField;
            }
            set
            {
                this.vacatedStyleField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string tag
        {
            get
            {
                return this.tagField;
            }
            set
            {
                this.tagField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte updatedVersion
        {
            get
            {
                return this.updatedVersionField;
            }
            set
            {
                this.updatedVersionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte minRefreshableVersion
        {
            get
            {
                return this.minRefreshableVersionField;
            }
            set
            {
                this.minRefreshableVersionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool asteriskTotals
        {
            get
            {
                return this.asteriskTotalsField;
            }
            set
            {
                this.asteriskTotalsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showItems
        {
            get
            {
                return this.showItemsField;
            }
            set
            {
                this.showItemsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool editData
        {
            get
            {
                return this.editDataField;
            }
            set
            {
                this.editDataField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool disableFieldList
        {
            get
            {
                return this.disableFieldListField;
            }
            set
            {
                this.disableFieldListField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showCalcMbrs
        {
            get
            {
                return this.showCalcMbrsField;
            }
            set
            {
                this.showCalcMbrsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool visualTotals
        {
            get
            {
                return this.visualTotalsField;
            }
            set
            {
                this.visualTotalsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showMultipleLabel
        {
            get
            {
                return this.showMultipleLabelField;
            }
            set
            {
                this.showMultipleLabelField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDataDropDown
        {
            get
            {
                return this.showDataDropDownField;
            }
            set
            {
                this.showDataDropDownField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDrill
        {
            get
            {
                return this.showDrillField;
            }
            set
            {
                this.showDrillField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool printDrill
        {
            get
            {
                return this.printDrillField;
            }
            set
            {
                this.printDrillField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showMemberPropertyTips
        {
            get
            {
                return this.showMemberPropertyTipsField;
            }
            set
            {
                this.showMemberPropertyTipsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDataTips
        {
            get
            {
                return this.showDataTipsField;
            }
            set
            {
                this.showDataTipsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool enableWizard
        {
            get
            {
                return this.enableWizardField;
            }
            set
            {
                this.enableWizardField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool enableDrill
        {
            get
            {
                return this.enableDrillField;
            }
            set
            {
                this.enableDrillField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool enableFieldProperties
        {
            get
            {
                return this.enableFieldPropertiesField;
            }
            set
            {
                this.enableFieldPropertiesField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool preserveFormatting
        {
            get
            {
                return this.preserveFormattingField;
            }
            set
            {
                this.preserveFormattingField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool useAutoFormatting
        {
            get
            {
                return this.useAutoFormattingField;
            }
            set
            {
                this.useAutoFormattingField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint pageWrap
        {
            get
            {
                return this.pageWrapField;
            }
            set
            {
                this.pageWrapField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool pageOverThenDown
        {
            get
            {
                return this.pageOverThenDownField;
            }
            set
            {
                this.pageOverThenDownField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool subtotalHiddenItems
        {
            get
            {
                return this.subtotalHiddenItemsField;
            }
            set
            {
                this.subtotalHiddenItemsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool rowGrandTotals
        {
            get
            {
                return this.rowGrandTotalsField;
            }
            set
            {
                this.rowGrandTotalsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool colGrandTotals
        {
            get
            {
                return this.colGrandTotalsField;
            }
            set
            {
                this.colGrandTotalsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool fieldPrintTitles
        {
            get
            {
                return this.fieldPrintTitlesField;
            }
            set
            {
                this.fieldPrintTitlesField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool itemPrintTitles
        {
            get
            {
                return this.itemPrintTitlesField;
            }
            set
            {
                this.itemPrintTitlesField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool mergeItem
        {
            get
            {
                return this.mergeItemField;
            }
            set
            {
                this.mergeItemField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDropZones
        {
            get
            {
                return this.showDropZonesField;
            }
            set
            {
                this.showDropZonesField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte createdVersion
        {
            get
            {
                return this.createdVersionField;
            }
            set
            {
                this.createdVersionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1")]
        public uint indent
        {
            get
            {
                return this.indentField;
            }
            set
            {
                this.indentField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showEmptyRow
        {
            get
            {
                return this.showEmptyRowField;
            }
            set
            {
                this.showEmptyRowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showEmptyCol
        {
            get
            {
                return this.showEmptyColField;
            }
            set
            {
                this.showEmptyColField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showHeaders
        {
            get
            {
                return this.showHeadersField;
            }
            set
            {
                this.showHeadersField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool compact
        {
            get
            {
                return this.compactField;
            }
            set
            {
                this.compactField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool outlineData
        {
            get
            {
                return this.outlineDataField;
            }
            set
            {
                this.outlineDataField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool compactData
        {
            get
            {
                return this.compactDataField;
            }
            set
            {
                this.compactDataField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool published
        {
            get
            {
                return this.publishedField;
            }
            set
            {
                this.publishedField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool gridDropZones
        {
            get
            {
                return this.gridDropZonesField;
            }
            set
            {
                this.gridDropZonesField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool immersive
        {
            get
            {
                return this.immersiveField;
            }
            set
            {
                this.immersiveField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool multipleFieldFilters
        {
            get
            {
                return this.multipleFieldFiltersField;
            }
            set
            {
                this.multipleFieldFiltersField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint chartFormat
        {
            get
            {
                return this.chartFormatField;
            }
            set
            {
                this.chartFormatField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string rowHeaderCaption
        {
            get
            {
                return this.rowHeaderCaptionField;
            }
            set
            {
                this.rowHeaderCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string colHeaderCaption
        {
            get
            {
                return this.colHeaderCaptionField;
            }
            set
            {
                this.colHeaderCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool fieldListSortAscending
        {
            get
            {
                return this.fieldListSortAscendingField;
            }
            set
            {
                this.fieldListSortAscendingField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool mdxSubqueries
        {
            get
            {
                return this.mdxSubqueriesField;
            }
            set
            {
                this.mdxSubqueriesField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool customListSort
        {
            get
            {
                return this.customListSortField;
            }
            set
            {
                this.customListSortField = value;
            }
        }

        public static CT_PivotTableDefinition Parse(System.IO.Stream is1)
        {
            throw new NotImplementedException();
        }

        public void Save(System.IO.Stream out1)
        {
            throw new NotImplementedException();
        }

        public CT_PivotTableStyle AddNewPivotTableStyleInfo()
        {
            throw new NotImplementedException();
        }

        public CT_RowFields AddNewRowFields()
        {
            throw new NotImplementedException();
        }

        public CT_ColFields AddNewColFields()
        {
            throw new NotImplementedException();
        }

        public CT_DataFields AddNewDataFields()
        {
            throw new NotImplementedException();
        }

        public CT_PageFields AddNewPageFields()
        {
            throw new NotImplementedException();
        }

        public CT_PivotFields AddNewPivotFields()
        {
            throw new NotImplementedException();
        }

        public CT_Location AddNewLocation()
        {
            throw new NotImplementedException();
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Location
    {

        private string refField;

        private uint firstHeaderRowField;

        private uint firstDataRowField;

        private uint firstDataColField;

        private uint rowPageCountField;

        private uint colPageCountField;

        public CT_Location()
        {
            this.rowPageCountField = ((uint)(0));
            this.colPageCountField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string @ref
        {
            get
            {
                return this.refField;
            }
            set
            {
                this.refField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint firstHeaderRow
        {
            get
            {
                return this.firstHeaderRowField;
            }
            set
            {
                this.firstHeaderRowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint firstDataRow
        {
            get
            {
                return this.firstDataRowField;
            }
            set
            {
                this.firstDataRowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint firstDataCol
        {
            get
            {
                return this.firstDataColField;
            }
            set
            {
                this.firstDataColField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint rowPageCount
        {
            get
            {
                return this.rowPageCountField;
            }
            set
            {
                this.rowPageCountField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint colPageCount
        {
            get
            {
                return this.colPageCountField;
            }
            set
            {
                this.colPageCountField = value;
            }
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotFields
    {

        private List<CT_PivotField> pivotFieldField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_PivotFields()
        {
            this.pivotFieldField = new List<CT_PivotField>();
        }

        [System.Xml.Serialization.XmlElementAttribute("pivotField", Order = 0)]
        public List<CT_PivotField> pivotField
        {
            get
            {
                return this.pivotFieldField;
            }
            set
            {
                this.pivotFieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        public void SetPivotFieldArray(int columnIndex, CT_PivotField pivotField)
        {
            throw new NotImplementedException();
        }

        public CT_PivotField AddNewPivotField()
        {
            throw new NotImplementedException();
        }

        public uint SizeOfPivotFieldArray()
        {
            throw new NotImplementedException();
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotField
    {

        private CT_Items itemsField;

        private CT_AutoSortScope autoSortScopeField;

        private CT_ExtensionList extLstField;

        private string nameField;

        private ST_Axis axisField;

        private bool axisFieldSpecified;

        private bool dataFieldField;

        private string subtotalCaptionField;

        private bool showDropDownsField;

        private bool hiddenLevelField;

        private string uniqueMemberPropertyField;

        private bool compactField;

        private bool allDrilledField;

        private uint numFmtIdField;

        private bool numFmtIdFieldSpecified;

        private bool outlineField;

        private bool subtotalTopField;

        private bool dragToRowField;

        private bool dragToColField;

        private bool multipleItemSelectionAllowedField;

        private bool dragToPageField;

        private bool dragToDataField;

        private bool dragOffField;

        private bool showAllField;

        private bool insertBlankRowField;

        private bool serverFieldField;

        private bool insertPageBreakField;

        private bool autoShowField;

        private bool topAutoShowField;

        private bool hideNewItemsField;

        private bool measureFilterField;

        private bool includeNewItemsInFilterField;

        private uint itemPageCountField;

        private ST_FieldSortType sortTypeField;

        private bool dataSourceSortField;

        private bool dataSourceSortFieldSpecified;

        private bool nonAutoSortDefaultField;

        private uint rankByField;

        private bool rankByFieldSpecified;

        private bool defaultSubtotalField;

        private bool sumSubtotalField;

        private bool countASubtotalField;

        private bool avgSubtotalField;

        private bool maxSubtotalField;

        private bool minSubtotalField;

        private bool productSubtotalField;

        private bool countSubtotalField;

        private bool stdDevSubtotalField;

        private bool stdDevPSubtotalField;

        private bool varSubtotalField;

        private bool varPSubtotalField;

        private bool showPropCellField;

        private bool showPropTipField;

        private bool showPropAsCaptionField;

        private bool defaultAttributeDrillStateField;

        public CT_PivotField()
        {
            this.extLstField = new CT_ExtensionList();
            this.autoSortScopeField = new CT_AutoSortScope();
            this.itemsField = new CT_Items();
            this.dataFieldField = false;
            this.showDropDownsField = true;
            this.hiddenLevelField = false;
            this.compactField = true;
            this.allDrilledField = false;
            this.outlineField = true;
            this.subtotalTopField = true;
            this.dragToRowField = true;
            this.dragToColField = true;
            this.multipleItemSelectionAllowedField = false;
            this.dragToPageField = true;
            this.dragToDataField = true;
            this.dragOffField = true;
            this.showAllField = true;
            this.insertBlankRowField = false;
            this.serverFieldField = false;
            this.insertPageBreakField = false;
            this.autoShowField = false;
            this.topAutoShowField = true;
            this.hideNewItemsField = false;
            this.measureFilterField = false;
            this.includeNewItemsInFilterField = false;
            this.itemPageCountField = ((uint)(10));
            this.sortTypeField = ST_FieldSortType.manual;
            this.nonAutoSortDefaultField = false;
            this.defaultSubtotalField = true;
            this.sumSubtotalField = false;
            this.countASubtotalField = false;
            this.avgSubtotalField = false;
            this.maxSubtotalField = false;
            this.minSubtotalField = false;
            this.productSubtotalField = false;
            this.countSubtotalField = false;
            this.stdDevSubtotalField = false;
            this.stdDevPSubtotalField = false;
            this.varSubtotalField = false;
            this.varPSubtotalField = false;
            this.showPropCellField = false;
            this.showPropTipField = false;
            this.showPropAsCaptionField = false;
            this.defaultAttributeDrillStateField = false;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_Items items
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

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_AutoSortScope autoSortScope
        {
            get
            {
                return this.autoSortScopeField;
            }
            set
            {
                this.autoSortScopeField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_Axis axis
        {
            get
            {
                return this.axisField;
            }
            set
            {
                this.axisField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool axisSpecified
        {
            get
            {
                return this.axisFieldSpecified;
            }
            set
            {
                this.axisFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool dataField
        {
            get
            {
                return this.dataFieldField;
            }
            set
            {
                this.dataFieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string subtotalCaption
        {
            get
            {
                return this.subtotalCaptionField;
            }
            set
            {
                this.subtotalCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDropDowns
        {
            get
            {
                return this.showDropDownsField;
            }
            set
            {
                this.showDropDownsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool hiddenLevel
        {
            get
            {
                return this.hiddenLevelField;
            }
            set
            {
                this.hiddenLevelField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string uniqueMemberProperty
        {
            get
            {
                return this.uniqueMemberPropertyField;
            }
            set
            {
                this.uniqueMemberPropertyField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool compact
        {
            get
            {
                return this.compactField;
            }
            set
            {
                this.compactField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool allDrilled
        {
            get
            {
                return this.allDrilledField;
            }
            set
            {
                this.allDrilledField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint numFmtId
        {
            get
            {
                return this.numFmtIdField;
            }
            set
            {
                this.numFmtIdField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool numFmtIdSpecified
        {
            get
            {
                return this.numFmtIdFieldSpecified;
            }
            set
            {
                this.numFmtIdFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool subtotalTop
        {
            get
            {
                return this.subtotalTopField;
            }
            set
            {
                this.subtotalTopField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToRow
        {
            get
            {
                return this.dragToRowField;
            }
            set
            {
                this.dragToRowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToCol
        {
            get
            {
                return this.dragToColField;
            }
            set
            {
                this.dragToColField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool multipleItemSelectionAllowed
        {
            get
            {
                return this.multipleItemSelectionAllowedField;
            }
            set
            {
                this.multipleItemSelectionAllowedField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToPage
        {
            get
            {
                return this.dragToPageField;
            }
            set
            {
                this.dragToPageField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToData
        {
            get
            {
                return this.dragToDataField;
            }
            set
            {
                this.dragToDataField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragOff
        {
            get
            {
                return this.dragOffField;
            }
            set
            {
                this.dragOffField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showAll
        {
            get
            {
                return this.showAllField;
            }
            set
            {
                this.showAllField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool insertBlankRow
        {
            get
            {
                return this.insertBlankRowField;
            }
            set
            {
                this.insertBlankRowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool serverField
        {
            get
            {
                return this.serverFieldField;
            }
            set
            {
                this.serverFieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool insertPageBreak
        {
            get
            {
                return this.insertPageBreakField;
            }
            set
            {
                this.insertPageBreakField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool autoShow
        {
            get
            {
                return this.autoShowField;
            }
            set
            {
                this.autoShowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool topAutoShow
        {
            get
            {
                return this.topAutoShowField;
            }
            set
            {
                this.topAutoShowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool hideNewItems
        {
            get
            {
                return this.hideNewItemsField;
            }
            set
            {
                this.hideNewItemsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool measureFilter
        {
            get
            {
                return this.measureFilterField;
            }
            set
            {
                this.measureFilterField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool includeNewItemsInFilter
        {
            get
            {
                return this.includeNewItemsInFilterField;
            }
            set
            {
                this.includeNewItemsInFilterField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "10")]
        public uint itemPageCount
        {
            get
            {
                return this.itemPageCountField;
            }
            set
            {
                this.itemPageCountField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_FieldSortType.manual)]
        public ST_FieldSortType sortType
        {
            get
            {
                return this.sortTypeField;
            }
            set
            {
                this.sortTypeField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool dataSourceSort
        {
            get
            {
                return this.dataSourceSortField;
            }
            set
            {
                this.dataSourceSortField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool dataSourceSortSpecified
        {
            get
            {
                return this.dataSourceSortFieldSpecified;
            }
            set
            {
                this.dataSourceSortFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool nonAutoSortDefault
        {
            get
            {
                return this.nonAutoSortDefaultField;
            }
            set
            {
                this.nonAutoSortDefaultField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint rankBy
        {
            get
            {
                return this.rankByField;
            }
            set
            {
                this.rankByField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool rankBySpecified
        {
            get
            {
                return this.rankByFieldSpecified;
            }
            set
            {
                this.rankByFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool defaultSubtotal
        {
            get
            {
                return this.defaultSubtotalField;
            }
            set
            {
                this.defaultSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool sumSubtotal
        {
            get
            {
                return this.sumSubtotalField;
            }
            set
            {
                this.sumSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool countASubtotal
        {
            get
            {
                return this.countASubtotalField;
            }
            set
            {
                this.countASubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool avgSubtotal
        {
            get
            {
                return this.avgSubtotalField;
            }
            set
            {
                this.avgSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool maxSubtotal
        {
            get
            {
                return this.maxSubtotalField;
            }
            set
            {
                this.maxSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool minSubtotal
        {
            get
            {
                return this.minSubtotalField;
            }
            set
            {
                this.minSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool productSubtotal
        {
            get
            {
                return this.productSubtotalField;
            }
            set
            {
                this.productSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool countSubtotal
        {
            get
            {
                return this.countSubtotalField;
            }
            set
            {
                this.countSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool stdDevSubtotal
        {
            get
            {
                return this.stdDevSubtotalField;
            }
            set
            {
                this.stdDevSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool stdDevPSubtotal
        {
            get
            {
                return this.stdDevPSubtotalField;
            }
            set
            {
                this.stdDevPSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool varSubtotal
        {
            get
            {
                return this.varSubtotalField;
            }
            set
            {
                this.varSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool varPSubtotal
        {
            get
            {
                return this.varPSubtotalField;
            }
            set
            {
                this.varPSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showPropCell
        {
            get
            {
                return this.showPropCellField;
            }
            set
            {
                this.showPropCellField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showPropTip
        {
            get
            {
                return this.showPropTipField;
            }
            set
            {
                this.showPropTipField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showPropAsCaption
        {
            get
            {
                return this.showPropAsCaptionField;
            }
            set
            {
                this.showPropAsCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool defaultAttributeDrillState
        {
            get
            {
                return this.defaultAttributeDrillStateField;
            }
            set
            {
                this.defaultAttributeDrillStateField = value;
            }
        }

        public CT_Items AddNewItems()
        {
            throw new NotImplementedException();
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Items
    {

        private List<CT_Item> itemField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_Items()
        {
            this.itemField = new List<CT_Item>();
        }

        [System.Xml.Serialization.XmlElementAttribute("item", Order = 0)]
        public List<CT_Item> item
        {
            get
            {
                return this.itemField;
            }
            set
            {
                this.itemField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        public CT_Item AddNewItem()
        {
            throw new NotImplementedException();
        }

        public uint SizeOfItemArray()
        {
            throw new NotImplementedException();
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Item
    {

        private string nField;

        private ST_ItemType tField;

        private bool hField;

        private bool sField;

        private bool sdField;

        private bool fField;

        private bool mField;

        private bool cField;

        private uint xField;

        private bool xFieldSpecified;

        private bool dField;

        private bool eField;

        public CT_Item()
        {
            this.tField = ST_ItemType.data;
            this.hField = false;
            this.sField = false;
            this.sdField = true;
            this.fField = false;
            this.mField = false;
            this.cField = false;
            this.dField = false;
            this.eField = true;
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string n
        {
            get
            {
                return this.nField;
            }
            set
            {
                this.nField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_ItemType.data)]
        public ST_ItemType t
        {
            get
            {
                return this.tField;
            }
            set
            {
                this.tField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool h
        {
            get
            {
                return this.hField;
            }
            set
            {
                this.hField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool s
        {
            get
            {
                return this.sField;
            }
            set
            {
                this.sField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool sd
        {
            get
            {
                return this.sdField;
            }
            set
            {
                this.sdField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool f
        {
            get
            {
                return this.fField;
            }
            set
            {
                this.fField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool m
        {
            get
            {
                return this.mField;
            }
            set
            {
                this.mField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool c
        {
            get
            {
                return this.cField;
            }
            set
            {
                this.cField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint x
        {
            get
            {
                return this.xField;
            }
            set
            {
                this.xField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool xSpecified
        {
            get
            {
                return this.xFieldSpecified;
            }
            set
            {
                this.xFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool d
        {
            get
            {
                return this.dField;
            }
            set
            {
                this.dField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool e
        {
            get
            {
                return this.eField;
            }
            set
            {
                this.eField = value;
            }
        }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_ItemType
    {

        /// <remarks/>
        data,

        /// <remarks/>
        @default,

        /// <remarks/>
        sum,

        /// <remarks/>
        countA,

        /// <remarks/>
        avg,

        /// <remarks/>
        max,

        /// <remarks/>
        min,

        /// <remarks/>
        product,

        /// <remarks/>
        count,

        /// <remarks/>
        stdDev,

        /// <remarks/>
        stdDevP,

        /// <remarks/>
        var,

        /// <remarks/>
        varP,

        /// <remarks/>
        grand,

        /// <remarks/>
        blank,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_AutoSortScope
    {

        private CT_PivotArea pivotAreaField;

        public CT_AutoSortScope()
        {
            this.pivotAreaField = new CT_PivotArea();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_PivotArea pivotArea
        {
            get
            {
                return this.pivotAreaField;
            }
            set
            {
                this.pivotAreaField = value;
            }
        }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_FieldSortType
    {

        /// <remarks/>
        manual,

        /// <remarks/>
        ascending,

        /// <remarks/>
        descending,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_RowFields
    {

        private List<CT_Field> fieldField;

        private uint countField;

        public CT_RowFields()
        {
            this.fieldField = new List<CT_Field>();
            this.countField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute("field", Order = 0)]
        public List<CT_Field> field
        {
            get
            {
                return this.fieldField;
            }
            set
            {
                this.fieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
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

        public CT_Field AddNewField()
        {
            throw new NotImplementedException();
        }

        public uint SizeOfFieldArray()
        {
            throw new NotImplementedException();
        }

        public IEnumerable<CT_Field> GetFieldArray()
        {
            throw new NotImplementedException();
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Field
    {

        private int xField;

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int x
        {
            get
            {
                return this.xField;
            }
            set
            {
                this.xField = value;
            }
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_rowItems
    {

        private List<CT_I> iField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_rowItems()
        {
            this.iField = new List<CT_I>();
        }

        [System.Xml.Serialization.XmlElementAttribute("i", Order = 0)]
        public List<CT_I> i
        {
            get
            {
                return this.iField;
            }
            set
            {
                this.iField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_I
    {

        private List<CT_X> xField;

        private ST_ItemType tField;

        private uint rField;

        private uint iField;

        public CT_I()
        {
            this.xField = new List<CT_X>();
            this.tField = ST_ItemType.data;
            this.rField = ((uint)(0));
            this.iField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute("x", Order = 0)]
        public List<CT_X> x
        {
            get
            {
                return this.xField;
            }
            set
            {
                this.xField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_ItemType.data)]
        public ST_ItemType t
        {
            get
            {
                return this.tField;
            }
            set
            {
                this.tField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint r
        {
            get
            {
                return this.rField;
            }
            set
            {
                this.rField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint i
        {
            get
            {
                return this.iField;
            }
            set
            {
                this.iField = value;
            }
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ColFields
    {

        private List<CT_Field> fieldField;

        private uint countField;

        public CT_ColFields()
        {
            this.fieldField = new List<CT_Field>();
            this.countField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute("field", Order = 0)]
        public List<CT_Field> field
        {
            get
            {
                return this.fieldField;
            }
            set
            {
                this.fieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
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

        public uint SizeOfFieldArray()
        {
            throw new NotImplementedException();
        }

        public CT_Field AddNewField()
        {
            throw new NotImplementedException();
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_colItems
    {

        private List<CT_I> iField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_colItems()
        {
            this.iField = new List<CT_I>();
        }

        [System.Xml.Serialization.XmlElementAttribute("i", Order = 0)]
        public List<CT_I> i
        {
            get
            {
                return this.iField;
            }
            set
            {
                this.iField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PageFields
    {

        private List<CT_PageField> pageFieldField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_PageFields()
        {
            this.pageFieldField = new List<CT_PageField>();
        }

        [System.Xml.Serialization.XmlElementAttribute("pageField", Order = 0)]
        public List<CT_PageField> pageField
        {
            get
            {
                return this.pageFieldField;
            }
            set
            {
                this.pageFieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        public CT_PageField AddNewPageField()
        {
            throw new NotImplementedException();
        }

        public uint SizeOfPageFieldArray()
        {
            throw new NotImplementedException();
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PageField
    {

        private CT_ExtensionList extLstField;

        private int fldField;

        private uint itemField;

        private bool itemFieldSpecified;

        private int hierField;

        private bool hierFieldSpecified;

        private string nameField;

        private string capField;

        public CT_PageField()
        {
            this.extLstField = new CT_ExtensionList();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int fld
        {
            get
            {
                return this.fldField;
            }
            set
            {
                this.fldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint item
        {
            get
            {
                return this.itemField;
            }
            set
            {
                this.itemField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool itemSpecified
        {
            get
            {
                return this.itemFieldSpecified;
            }
            set
            {
                this.itemFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int hier
        {
            get
            {
                return this.hierField;
            }
            set
            {
                this.hierField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool hierSpecified
        {
            get
            {
                return this.hierFieldSpecified;
            }
            set
            {
                this.hierFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string cap
        {
            get
            {
                return this.capField;
            }
            set
            {
                this.capField = value;
            }
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_DataFields
    {

        private List<CT_DataField> dataFieldField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_DataFields()
        {
            this.dataFieldField = new List<CT_DataField>();
        }

        [System.Xml.Serialization.XmlElementAttribute("dataField", Order = 0)]
        public List<CT_DataField> dataField
        {
            get
            {
                return this.dataFieldField;
            }
            set
            {
                this.dataFieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        public CT_DataField AddNewDataField()
        {
            throw new NotImplementedException();
        }

        public uint SizeOfDataFieldArray()
        {
            throw new NotImplementedException();
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_DataField
    {

        private CT_ExtensionList extLstField;

        private string nameField;

        private uint fldField;

        private ST_DataConsolidateFunction subtotalField;

        private ST_ShowDataAs showDataAsField;

        private int baseFieldField;

        private uint baseItemField;

        private uint numFmtIdField;

        private bool numFmtIdFieldSpecified;

        public CT_DataField()
        {
            this.extLstField = new CT_ExtensionList();
            this.subtotalField = ST_DataConsolidateFunction.sum;
            this.showDataAsField = ST_ShowDataAs.normal;
            this.baseFieldField = -1;
            this.baseItemField = ((uint)(1048832));
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint fld
        {
            get
            {
                return this.fldField;
            }
            set
            {
                this.fldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_DataConsolidateFunction.sum)]
        public ST_DataConsolidateFunction subtotal
        {
            get
            {
                return this.subtotalField;
            }
            set
            {
                this.subtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_ShowDataAs.normal)]
        public ST_ShowDataAs showDataAs
        {
            get
            {
                return this.showDataAsField;
            }
            set
            {
                this.showDataAsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(-1)]
        public int baseField
        {
            get
            {
                return this.baseFieldField;
            }
            set
            {
                this.baseFieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1048832")]
        public uint baseItem
        {
            get
            {
                return this.baseItemField;
            }
            set
            {
                this.baseItemField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint numFmtId
        {
            get
            {
                return this.numFmtIdField;
            }
            set
            {
                this.numFmtIdField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool numFmtIdSpecified
        {
            get
            {
                return this.numFmtIdFieldSpecified;
            }
            set
            {
                this.numFmtIdFieldSpecified = value;
            }
        }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_ShowDataAs
    {

        /// <remarks/>
        normal,

        /// <remarks/>
        difference,

        /// <remarks/>
        percent,

        /// <remarks/>
        percentDiff,

        /// <remarks/>
        runTotal,

        /// <remarks/>
        percentOfRow,

        /// <remarks/>
        percentOfCol,

        /// <remarks/>
        percentOfTotal,

        /// <remarks/>
        index,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Formats
    {

        private List<CT_Format> formatField;

        private uint countField;

        public CT_Formats()
        {
            this.formatField = new List<CT_Format>();
            this.countField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute("format", Order = 0)]
        public List<CT_Format> format
        {
            get
            {
                return this.formatField;
            }
            set
            {
                this.formatField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
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
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Format
    {

        private CT_PivotArea pivotAreaField;

        private CT_ExtensionList extLstField;

        private ST_FormatAction actionField;

        private uint dxfIdField;

        private bool dxfIdFieldSpecified;

        public CT_Format()
        {
            this.extLstField = new CT_ExtensionList();
            this.pivotAreaField = new CT_PivotArea();
            this.actionField = ST_FormatAction.formatting;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_PivotArea pivotArea
        {
            get
            {
                return this.pivotAreaField;
            }
            set
            {
                this.pivotAreaField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_FormatAction.formatting)]
        public ST_FormatAction action
        {
            get
            {
                return this.actionField;
            }
            set
            {
                this.actionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_FormatAction
    {

        /// <remarks/>
        blank,

        /// <remarks/>
        formatting,

        /// <remarks/>
        drill,

        /// <remarks/>
        formula,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ConditionalFormats
    {

        private List<CT_ConditionalFormat> conditionalFormatField;

        private uint countField;

        public CT_ConditionalFormats()
        {
            this.conditionalFormatField = new List<CT_ConditionalFormat>();
            this.countField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute("conditionalFormat", Order = 0)]
        public List<CT_ConditionalFormat> conditionalFormat
        {
            get
            {
                return this.conditionalFormatField;
            }
            set
            {
                this.conditionalFormatField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
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
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ConditionalFormat
    {

        private CT_PivotAreas pivotAreasField;

        private CT_ExtensionList extLstField;

        private ST_Scope scopeField;

        private ST_Type typeField;

        private uint priorityField;

        public CT_ConditionalFormat()
        {
            this.extLstField = new CT_ExtensionList();
            this.pivotAreasField = new CT_PivotAreas();
            this.scopeField = ST_Scope.selection;
            this.typeField = ST_Type.none;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_PivotAreas pivotAreas
        {
            get
            {
                return this.pivotAreasField;
            }
            set
            {
                this.pivotAreasField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_Scope.selection)]
        public ST_Scope scope
        {
            get
            {
                return this.scopeField;
            }
            set
            {
                this.scopeField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_Type.none)]
        public ST_Type type
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint priority
        {
            get
            {
                return this.priorityField;
            }
            set
            {
                this.priorityField = value;
            }
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotAreas
    {

        private List<CT_PivotArea> pivotAreaField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_PivotAreas()
        {
            this.pivotAreaField = new List<CT_PivotArea>();
        }

        [System.Xml.Serialization.XmlElementAttribute("pivotArea", Order = 0)]
        public List<CT_PivotArea> pivotArea
        {
            get
            {
                return this.pivotAreaField;
            }
            set
            {
                this.pivotAreaField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_Scope
    {

        /// <remarks/>
        selection,

        /// <remarks/>
        data,

        /// <remarks/>
        field,
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_Type
    {

        /// <remarks/>
        none,

        /// <remarks/>
        all,

        /// <remarks/>
        row,

        /// <remarks/>
        column,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ChartFormats
    {

        private List<CT_ChartFormat> chartFormatField;

        private uint countField;

        public CT_ChartFormats()
        {
            this.chartFormatField = new List<CT_ChartFormat>();
            this.countField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute("chartFormat", Order = 0)]
        public List<CT_ChartFormat> chartFormat
        {
            get
            {
                return this.chartFormatField;
            }
            set
            {
                this.chartFormatField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
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
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ChartFormat
    {

        private CT_PivotArea pivotAreaField;

        private uint chartField;

        private uint formatField;

        private bool seriesField;

        public CT_ChartFormat()
        {
            this.pivotAreaField = new CT_PivotArea();
            this.seriesField = false;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_PivotArea pivotArea
        {
            get
            {
                return this.pivotAreaField;
            }
            set
            {
                this.pivotAreaField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint chart
        {
            get
            {
                return this.chartField;
            }
            set
            {
                this.chartField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint format
        {
            get
            {
                return this.formatField;
            }
            set
            {
                this.formatField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool series
        {
            get
            {
                return this.seriesField;
            }
            set
            {
                this.seriesField = value;
            }
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotHierarchies
    {

        private List<CT_PivotHierarchy> pivotHierarchyField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_PivotHierarchies()
        {
            this.pivotHierarchyField = new List<CT_PivotHierarchy>();
        }

        [System.Xml.Serialization.XmlElementAttribute("pivotHierarchy", Order = 0)]
        public List<CT_PivotHierarchy> pivotHierarchy
        {
            get
            {
                return this.pivotHierarchyField;
            }
            set
            {
                this.pivotHierarchyField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotHierarchy
    {

        private CT_MemberProperties mpsField;

        private List<CT_Members> membersField;

        private CT_ExtensionList extLstField;

        private bool outlineField;

        private bool multipleItemSelectionAllowedField;

        private bool subtotalTopField;

        private bool showInFieldListField;

        private bool dragToRowField;

        private bool dragToColField;

        private bool dragToPageField;

        private bool dragToDataField;

        private bool dragOffField;

        private bool includeNewItemsInFilterField;

        private string captionField;

        public CT_PivotHierarchy()
        {
            this.extLstField = new CT_ExtensionList();
            this.membersField = new List<CT_Members>();
            this.mpsField = new CT_MemberProperties();
            this.outlineField = false;
            this.multipleItemSelectionAllowedField = false;
            this.subtotalTopField = false;
            this.showInFieldListField = true;
            this.dragToRowField = true;
            this.dragToColField = true;
            this.dragToPageField = true;
            this.dragToDataField = false;
            this.dragOffField = true;
            this.includeNewItemsInFilterField = false;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_MemberProperties mps
        {
            get
            {
                return this.mpsField;
            }
            set
            {
                this.mpsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute("members", Order = 1)]
        public List<CT_Members> members
        {
            get
            {
                return this.membersField;
            }
            set
            {
                this.membersField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool multipleItemSelectionAllowed
        {
            get
            {
                return this.multipleItemSelectionAllowedField;
            }
            set
            {
                this.multipleItemSelectionAllowedField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool subtotalTop
        {
            get
            {
                return this.subtotalTopField;
            }
            set
            {
                this.subtotalTopField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showInFieldList
        {
            get
            {
                return this.showInFieldListField;
            }
            set
            {
                this.showInFieldListField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToRow
        {
            get
            {
                return this.dragToRowField;
            }
            set
            {
                this.dragToRowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToCol
        {
            get
            {
                return this.dragToColField;
            }
            set
            {
                this.dragToColField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToPage
        {
            get
            {
                return this.dragToPageField;
            }
            set
            {
                this.dragToPageField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool dragToData
        {
            get
            {
                return this.dragToDataField;
            }
            set
            {
                this.dragToDataField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragOff
        {
            get
            {
                return this.dragOffField;
            }
            set
            {
                this.dragOffField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool includeNewItemsInFilter
        {
            get
            {
                return this.includeNewItemsInFilterField;
            }
            set
            {
                this.includeNewItemsInFilterField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string caption
        {
            get
            {
                return this.captionField;
            }
            set
            {
                this.captionField = value;
            }
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_MemberProperties
    {

        private List<CT_MemberProperty> mpField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_MemberProperties()
        {
            this.mpField = new List<CT_MemberProperty>();
        }

        [System.Xml.Serialization.XmlElementAttribute("mp", Order = 0)]
        public List<CT_MemberProperty> mp
        {
            get
            {
                return this.mpField;
            }
            set
            {
                this.mpField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_MemberProperty
    {

        private string nameField;

        private bool showCellField;

        private bool showTipField;

        private bool showAsCaptionField;

        private uint nameLenField;

        private bool nameLenFieldSpecified;

        private uint pPosField;

        private bool pPosFieldSpecified;

        private uint pLenField;

        private bool pLenFieldSpecified;

        private uint levelField;

        private bool levelFieldSpecified;

        private uint fieldField;

        public CT_MemberProperty()
        {
            this.showCellField = false;
            this.showTipField = false;
            this.showAsCaptionField = false;
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showCell
        {
            get
            {
                return this.showCellField;
            }
            set
            {
                this.showCellField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showTip
        {
            get
            {
                return this.showTipField;
            }
            set
            {
                this.showTipField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showAsCaption
        {
            get
            {
                return this.showAsCaptionField;
            }
            set
            {
                this.showAsCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint nameLen
        {
            get
            {
                return this.nameLenField;
            }
            set
            {
                this.nameLenField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool nameLenSpecified
        {
            get
            {
                return this.nameLenFieldSpecified;
            }
            set
            {
                this.nameLenFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint pPos
        {
            get
            {
                return this.pPosField;
            }
            set
            {
                this.pPosField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool pPosSpecified
        {
            get
            {
                return this.pPosFieldSpecified;
            }
            set
            {
                this.pPosFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint pLen
        {
            get
            {
                return this.pLenField;
            }
            set
            {
                this.pLenField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool pLenSpecified
        {
            get
            {
                return this.pLenFieldSpecified;
            }
            set
            {
                this.pLenFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint level
        {
            get
            {
                return this.levelField;
            }
            set
            {
                this.levelField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool levelSpecified
        {
            get
            {
                return this.levelFieldSpecified;
            }
            set
            {
                this.levelFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint field
        {
            get
            {
                return this.fieldField;
            }
            set
            {
                this.fieldField = value;
            }
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Members
    {

        private List<CT_Member> memberField;

        private uint countField;

        private bool countFieldSpecified;

        private uint levelField;

        private bool levelFieldSpecified;

        public CT_Members()
        {
            this.memberField = new List<CT_Member>();
        }

        [System.Xml.Serialization.XmlElementAttribute("member", Order = 0)]
        public List<CT_Member> member
        {
            get
            {
                return this.memberField;
            }
            set
            {
                this.memberField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint level
        {
            get
            {
                return this.levelField;
            }
            set
            {
                this.levelField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool levelSpecified
        {
            get
            {
                return this.levelFieldSpecified;
            }
            set
            {
                this.levelFieldSpecified = value;
            }
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Member
    {

        private string nameField;

        [System.Xml.Serialization.XmlAttributeAttribute()]
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
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotTableStyle
    {

        private string nameField;

        private bool showRowHeadersField;

        private bool showRowHeadersFieldSpecified;

        private bool showColHeadersField;

        private bool showColHeadersFieldSpecified;

        private bool showRowStripesField;

        private bool showRowStripesFieldSpecified;

        private bool showColStripesField;

        private bool showColStripesFieldSpecified;

        private bool showLastColumnField;

        private bool showLastColumnFieldSpecified;

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showRowHeaders
        {
            get
            {
                return this.showRowHeadersField;
            }
            set
            {
                this.showRowHeadersField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showRowHeadersSpecified
        {
            get
            {
                return this.showRowHeadersFieldSpecified;
            }
            set
            {
                this.showRowHeadersFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showColHeaders
        {
            get
            {
                return this.showColHeadersField;
            }
            set
            {
                this.showColHeadersField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showColHeadersSpecified
        {
            get
            {
                return this.showColHeadersFieldSpecified;
            }
            set
            {
                this.showColHeadersFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showRowStripes
        {
            get
            {
                return this.showRowStripesField;
            }
            set
            {
                this.showRowStripesField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showRowStripesSpecified
        {
            get
            {
                return this.showRowStripesFieldSpecified;
            }
            set
            {
                this.showRowStripesFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showColStripes
        {
            get
            {
                return this.showColStripesField;
            }
            set
            {
                this.showColStripesField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showColStripesSpecified
        {
            get
            {
                return this.showColStripesFieldSpecified;
            }
            set
            {
                this.showColStripesFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showLastColumn
        {
            get
            {
                return this.showLastColumnField;
            }
            set
            {
                this.showLastColumnField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showLastColumnSpecified
        {
            get
            {
                return this.showLastColumnFieldSpecified;
            }
            set
            {
                this.showLastColumnFieldSpecified = value;
            }
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotFilters
    {

        private List<CT_PivotFilter> filterField;

        private uint countField;

        public CT_PivotFilters()
        {
            this.filterField = new List<CT_PivotFilter>();
            this.countField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute("filter", Order = 0)]
        public List<CT_PivotFilter> filter
        {
            get
            {
                return this.filterField;
            }
            set
            {
                this.filterField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
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
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotFilter
    {

        private CT_AutoFilter autoFilterField;

        private CT_ExtensionList extLstField;

        private uint fldField;

        private uint mpFldField;

        private bool mpFldFieldSpecified;

        private ST_PivotFilterType typeField;

        private int evalOrderField;

        private uint idField;

        private uint iMeasureHierField;

        private bool iMeasureHierFieldSpecified;

        private uint iMeasureFldField;

        private bool iMeasureFldFieldSpecified;

        private string nameField;

        private string descriptionField;

        private string stringValue1Field;

        private string stringValue2Field;

        public CT_PivotFilter()
        {
            this.extLstField = new CT_ExtensionList();
            this.autoFilterField = new CT_AutoFilter();
            this.evalOrderField = 0;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_AutoFilter autoFilter
        {
            get
            {
                return this.autoFilterField;
            }
            set
            {
                this.autoFilterField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint fld
        {
            get
            {
                return this.fldField;
            }
            set
            {
                this.fldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint mpFld
        {
            get
            {
                return this.mpFldField;
            }
            set
            {
                this.mpFldField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool mpFldSpecified
        {
            get
            {
                return this.mpFldFieldSpecified;
            }
            set
            {
                this.mpFldFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_PivotFilterType type
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public int evalOrder
        {
            get
            {
                return this.evalOrderField;
            }
            set
            {
                this.evalOrderField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint iMeasureHier
        {
            get
            {
                return this.iMeasureHierField;
            }
            set
            {
                this.iMeasureHierField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool iMeasureHierSpecified
        {
            get
            {
                return this.iMeasureHierFieldSpecified;
            }
            set
            {
                this.iMeasureHierFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint iMeasureFld
        {
            get
            {
                return this.iMeasureFldField;
            }
            set
            {
                this.iMeasureFldField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool iMeasureFldSpecified
        {
            get
            {
                return this.iMeasureFldFieldSpecified;
            }
            set
            {
                this.iMeasureFldFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string description
        {
            get
            {
                return this.descriptionField;
            }
            set
            {
                this.descriptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string stringValue1
        {
            get
            {
                return this.stringValue1Field;
            }
            set
            {
                this.stringValue1Field = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string stringValue2
        {
            get
            {
                return this.stringValue2Field;
            }
            set
            {
                this.stringValue2Field = value;
            }
        }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_PivotFilterType
    {

        /// <remarks/>
        unknown,

        /// <remarks/>
        count,

        /// <remarks/>
        percent,

        /// <remarks/>
        sum,

        /// <remarks/>
        captionEqual,

        /// <remarks/>
        captionNotEqual,

        /// <remarks/>
        captionBeginsWith,

        /// <remarks/>
        captionNotBeginsWith,

        /// <remarks/>
        captionEndsWith,

        /// <remarks/>
        captionNotEndsWith,

        /// <remarks/>
        captionContains,

        /// <remarks/>
        captionNotContains,

        /// <remarks/>
        captionGreaterThan,

        /// <remarks/>
        captionGreaterThanOrEqual,

        /// <remarks/>
        captionLessThan,

        /// <remarks/>
        captionLessThanOrEqual,

        /// <remarks/>
        captionBetween,

        /// <remarks/>
        captionNotBetween,

        /// <remarks/>
        valueEqual,

        /// <remarks/>
        valueNotEqual,

        /// <remarks/>
        valueGreaterThan,

        /// <remarks/>
        valueGreaterThanOrEqual,

        /// <remarks/>
        valueLessThan,

        /// <remarks/>
        valueLessThanOrEqual,

        /// <remarks/>
        valueBetween,

        /// <remarks/>
        valueNotBetween,

        /// <remarks/>
        dateEqual,

        /// <remarks/>
        dateNotEqual,

        /// <remarks/>
        dateOlderThan,

        /// <remarks/>
        dateOlderThanOrEqual,

        /// <remarks/>
        dateNewerThan,

        /// <remarks/>
        dateNewerThanOrEqual,

        /// <remarks/>
        dateBetween,

        /// <remarks/>
        dateNotBetween,

        /// <remarks/>
        tomorrow,

        /// <remarks/>
        today,

        /// <remarks/>
        yesterday,

        /// <remarks/>
        nextWeek,

        /// <remarks/>
        thisWeek,

        /// <remarks/>
        lastWeek,

        /// <remarks/>
        nextMonth,

        /// <remarks/>
        thisMonth,

        /// <remarks/>
        lastMonth,

        /// <remarks/>
        nextQuarter,

        /// <remarks/>
        thisQuarter,

        /// <remarks/>
        lastQuarter,

        /// <remarks/>
        nextYear,

        /// <remarks/>
        thisYear,

        /// <remarks/>
        lastYear,

        /// <remarks/>
        yearToDate,

        /// <remarks/>
        Q1,

        /// <remarks/>
        Q2,

        /// <remarks/>
        Q3,

        /// <remarks/>
        Q4,

        /// <remarks/>
        M1,

        /// <remarks/>
        M2,

        /// <remarks/>
        M3,

        /// <remarks/>
        M4,

        /// <remarks/>
        M5,

        /// <remarks/>
        M6,

        /// <remarks/>
        M7,

        /// <remarks/>
        M8,

        /// <remarks/>
        M9,

        /// <remarks/>
        M10,

        /// <remarks/>
        M11,

        /// <remarks/>
        M12,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_RowHierarchiesUsage
    {

        private List<CT_HierarchyUsage> rowHierarchyUsageField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_RowHierarchiesUsage()
        {
            this.rowHierarchyUsageField = new List<CT_HierarchyUsage>();
        }

        [System.Xml.Serialization.XmlElementAttribute("rowHierarchyUsage", Order = 0)]
        public List<CT_HierarchyUsage> rowHierarchyUsage
        {
            get
            {
                return this.rowHierarchyUsageField;
            }
            set
            {
                this.rowHierarchyUsageField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_HierarchyUsage
    {

        private int hierarchyUsageField;

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int hierarchyUsage
        {
            get
            {
                return this.hierarchyUsageField;
            }
            set
            {
                this.hierarchyUsageField = value;
            }
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ColHierarchiesUsage
    {

        private List<CT_HierarchyUsage> colHierarchyUsageField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_ColHierarchiesUsage()
        {
            this.colHierarchyUsageField = new List<CT_HierarchyUsage>();
        }

        [System.Xml.Serialization.XmlElementAttribute("colHierarchyUsage", Order = 0)]
        public List<CT_HierarchyUsage> colHierarchyUsage
        {
            get
            {
                return this.colHierarchyUsageField;
            }
            set
            {
                this.colHierarchyUsageField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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
