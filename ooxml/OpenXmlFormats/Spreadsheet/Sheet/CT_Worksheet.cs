using System;
using System.Collections.Generic;
using System.IO;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    //[System.Diagnostics.DebuggerStepThrough]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ElementName = "worksheet",
        IsNullable = false)]
    public class CT_Worksheet
    {
        // all the class attributes are XML elements. All except sheetData are optional.
        private CT_SheetPr sheetPrField = null;

        private CT_SheetDimension dimensionField = null;

        private CT_SheetViews sheetViewsField = null;

        private CT_SheetFormatPr sheetFormatPrField = null;

        private List<CT_Cols> colsField = null;

        private CT_SheetData sheetDataField = new CT_SheetData();

        private CT_SheetCalcPr sheetCalcPrField = null;

        private CT_SheetProtection sheetProtectionField = null;

        private List<CT_ProtectedRange> protectedRangesField = null;

        private CT_Scenarios scenariosField = null;

        private CT_AutoFilter autoFilterField = null;

        private CT_SortState sortStateField = null;

        private CT_DataConsolidate dataConsolidateField = null;

        private List<CT_CustomSheetView> customSheetViewsField = null;

        private CT_MergeCells mergeCellsField = null;

        private CT_PhoneticPr phoneticPrField = null;

        private List<CT_ConditionalFormatting> conditionalFormattingField = null;

        private CT_DataValidations dataValidationsField = null;

        private CT_Hyperlinks hyperlinksField = null;

        private CT_PrintOptions printOptionsField = null;

        private CT_PageMargins pageMarginsField = null;

        private CT_PageSetup pageSetupField = null;

        private CT_HeaderFooter headerFooterField = null;

        private CT_PageBreak rowBreaksField = null;

        private CT_PageBreak colBreaksField = null;

        private List<CT_CustomProperty> customPropertiesField = null;

        private List<CT_CellWatch> cellWatchesField = null;

        private CT_IgnoredErrors ignoredErrorsField = null;

        private List<CT_CellSmartTags> smartTagsField = null;

        private CT_Drawing drawingField = null;

        private CT_LegacyDrawing legacyDrawingField = null;

        private CT_LegacyDrawing legacyDrawingHFField = null;

        private CT_SheetBackgroundPicture pictureField = null;

        private List<CT_OleObject> oleObjectsField = null;

        private List<CT_Control> controlsField = null;

        private CT_WebPublishItems webPublishItemsField = null;

        private CT_TableParts tablePartsField = null;

        private CT_ExtensionList extLstField = null;


        public CT_AutoFilter AddNewAutoFilter()
        {
            this.autoFilterField = new CT_AutoFilter();
            return this.autoFilterField;
        }
        public void Save(Stream stream)
        {
            WorksheetDocument.serializer.Serialize(stream, this, WorksheetDocument.namespaces);
        }
        public bool IsSetRowBreaks()
        {
            return this.rowBreaksField != null;
        }
        public CT_Drawing AddNewDrawing()
        {
            this.drawingField = new CT_Drawing();
            return drawingField;
        }
        public CT_LegacyDrawing AddNewLegacyDrawing()
        {
            this.legacyDrawing = new CT_LegacyDrawing();
            return legacyDrawing;
        }
        public CT_PageBreak AddNewRowBreaks()
        {
            this.rowBreaksField = new CT_PageBreak();
            return this.rowBreaksField;
        }
        public CT_PageBreak AddNewColBreaks()
        {
            this.colBreaksField = new CT_PageBreak();
            return this.colBreaksField;
        }
        public bool IsSetSheetFormatPr()
        {
            return this.sheetFormatPrField != null;
        }
        public bool IsSetPrintOptions()
        {
            return this.printOptionsField != null;
        }
        public void UnsetMergeCells()
        {
            this.mergeCellsField = null;
        }
        public CT_PrintOptions AddNewPrintOptions()
        {
            this.printOptionsField = new CT_PrintOptions();
            return this.printOptionsField;
        }
        public CT_DataValidations AddNewDataValidations()
        {
            this.dataValidationsField = new CT_DataValidations();
            return this.dataValidationsField;
        }
        public CT_SheetViews AddNewSheetViews()
        {
            this.sheetViewsField = new CT_SheetViews();
            return this.sheetViewsField;
        }
        public CT_Hyperlinks AddNewHyperlinks()
        {
            this.hyperlinksField = new CT_Hyperlinks();
            return this.hyperlinksField;
        }
        public CT_ConditionalFormatting AddNewConditionalFormatting()
        {
            if (null == conditionalFormattingField) { conditionalFormattingField = new List<CT_ConditionalFormatting>(); }
            CT_ConditionalFormatting cf = new CT_ConditionalFormatting();
            this.conditionalFormattingField.Add(cf);
            return cf;
        }
        public CT_ConditionalFormatting GetConditionalFormattingArray(int index)
        {
            return (null == conditionalFormattingField) ? null : this.conditionalFormattingField[index];
        }
        public CT_MergeCells AddNewMergeCells()
        {
            this.mergeCellsField = new CT_MergeCells();
            return this.mergeCellsField;
        }
        public bool IsSetColBreaks()
        {
            return this.colBreaksField != null;
        }
        public bool IsSetHyperlinks()
        {
            return this.hyperlinksField != null;
        }
        public bool IsSetMergeCells()
        {
            return this.mergeCellsField != null;
        }
        public bool IsSetSheetProtection()
        {
            return this.sheetProtectionField != null;
        }
        public bool IsSetDrawing()
        {
            return this.drawingField != null;
        }
        public void UnsetDrawing()
        {
            this.drawingField = null;
        }
        public bool IsSetLegacyDrawing()
        {
            return this.legacyDrawingField != null;
        }
        public void UnsetLegacyDrawing()
        {
            this.legacyDrawingField = null;
        }
        public bool IsSetPageSetup()
        {
            return this.pageSetupField != null;
        }
        public bool IsSetTableParts()
        {
            return this.tablePartsField != null;
        }
        public bool IsSetSheetCalcPr()
        {
            return this.sheetCalcPrField != null;
        }
        public CT_SheetProtection AddNewSheetProtection()
        {
            this.sheetProtectionField = new CT_SheetProtection();
            return this.sheetProtectionField;
        }
        public CT_TableParts AddNewTableParts()
        {
            this.tablePartsField = new CT_TableParts();
            return this.tablePartsField;
        }
        public CT_PageMargins AddNewPageMargins()
        {
            this.pageMarginsField = new CT_PageMargins();
            return this.pageMarginsField;
        }
        public CT_PageSetup AddNewPageSetup()
        {
            this.pageSetupField = new CT_PageSetup();
            return this.pageSetupField;
        }
        public void SetColsArray(List<CT_Cols> a)
        {
            this.colsField = a;
        }
        public int sizeOfColsArray()
        {
            return (null == colsField) ? 0 : this.colsField.Count;
        }
        public void RemoveCols(int index)
        {
            this.colsField.RemoveAt(index);
        }
        public CT_Cols AddNewCols()
        {
            if (null == colsField) { colsField = new List<CT_Cols>(); }
            CT_Cols newCols = new CT_Cols();
            this.colsField.Add(newCols);
            return newCols;
        }
        public void SetColsArray(int index, CT_Cols newCols)
        {
            if (null == colsField)
            {
                colsField = new List<CT_Cols>();
            }
            else
            {
                colsField.Clear();
            }
            this.colsField.Insert(index, newCols);
        }
        public CT_Cols GetColsArray(int index)
        {
            if (null == colsField) { colsField = new List<CT_Cols>(); }
            return this.colsField[index];
        }
        public List<CT_Cols> GetColsArray()
        {
            return this.colsField;
        }
        public bool IsSetPageMargins()
        {
            return this.pageMarginsField != null;
        }
        public bool IsSetHyperLinks()
        {
            return this.hyperlinksField != null;
        }
        public bool IsSetSheetPr()
        {
            return this.sheetPrField != null;
        }
        public int SizeOfConditionalFormattingArray()
        {
            return (null == conditionalFormattingField) ? 0 : this.conditionalFormatting.Count;
        }

        public void UnsetSheetProtection()
        {
            this.sheetProtectionField = null;
        }

        public CT_SheetFormatPr AddNewSheetFormatPr()
        {
            this.sheetFormatPrField = new CT_SheetFormatPr();
            return sheetFormatPrField;
        }
        public CT_SheetCalcPr AddNewSheetCalcPr()
        {
            this.sheetCalcPrField = new CT_SheetCalcPr();
            return sheetCalcPrField;
        }
        public CT_SheetPr AddNewSheetPr()
        {
            this.sheetPrField = new CT_SheetPr();
            return sheetPrField;
        }
        public CT_SheetDimension AddNewDimension()
        {
            this.dimensionField = new CT_SheetDimension();
            return dimensionField;
        }
        public CT_SheetData AddNewSheetData()
        {
            this.sheetDataField = new CT_SheetData();
            return sheetDataField;
        }
        [XmlElement("sheetPr")]//, Order=0)]
        public CT_SheetPr sheetPr
        {
            get
            {
                return this.sheetPrField;
            }
            set
            {
                this.sheetPrField = value;
            }
        }

        [XmlElement]
        public CT_SheetDimension dimension
        {
            get
            {
                return this.dimensionField;
            }
            set
            {
                this.dimensionField = value;
            }
        }
        [XmlElement]
        public CT_SheetViews sheetViews
        {
            get
            {
                return this.sheetViewsField;
            }
            set
            {
                this.sheetViewsField = value;
            }
        }
        [XmlElement]
        public CT_SheetFormatPr sheetFormatPr
        {
            get
            {
                return this.sheetFormatPrField;
            }
            set
            {
                this.sheetFormatPrField = value;
            }
        }

        //[XmlArray(Order = 4)]
        // [XmlArrayItem("cols", typeof(CT_Cols), IsNullable = false)]
        [XmlElement("cols")]
        public List<CT_Cols> cols
        {
            get
            {
                return this.colsField;
            }
            set
            {
                this.colsField = value;
            }
        }

        [XmlElement("sheetData", IsNullable = false)]
        public CT_SheetData sheetData
        {
            get
            {
                return this.sheetDataField;
            }
            set
            {
                this.sheetDataField = value;
            }
        }
        [XmlElement]
        public CT_SheetCalcPr sheetCalcPr
        {
            get
            {
                return this.sheetCalcPrField;
            }
            set
            {
                this.sheetCalcPrField = value;
            }
        }
        [XmlElement]
        public CT_SheetProtection sheetProtection
        {
            get
            {
                return this.sheetProtectionField;
            }
            set
            {
                this.sheetProtectionField = value;
            }
        }

        //[XmlArray(Order = 8)]
        //[XmlArrayItem("protectedRange", IsNullable = false)]
        [XmlElement("protectedRange")]
        public List<CT_ProtectedRange> protectedRanges
        {
            get
            {
                return this.protectedRangesField;
            }
            set
            {
                this.protectedRangesField = value;
            }
        }

        public CT_Scenarios scenarios
        {
            get
            {
                return this.scenariosField;
            }
            set
            {
                this.scenariosField = value;
            }
        }

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

        public CT_SortState sortState
        {
            get
            {
                return this.sortStateField;
            }
            set
            {
                this.sortStateField = value;
            }
        }

        public CT_DataConsolidate dataConsolidate
        {
            get
            {
                return this.dataConsolidateField;
            }
            set
            {
                this.dataConsolidateField = value;
            }
        }

        //[XmlArray(Order = 13)]
        //[XmlArrayItem("customSheetView", IsNullable = false)]
        [XmlElement("customSheetView")]
        public List<CT_CustomSheetView> customSheetViews
        {
            get
            {
                return this.customSheetViewsField;
            }
            set
            {
                this.customSheetViewsField = value;
            }
        }
        [XmlElement]
        public CT_MergeCells mergeCells
        {
            get
            {
                return this.mergeCellsField;
            }
            set
            {
                this.mergeCellsField = value;
            }
        }
        [XmlElement]
        public CT_PhoneticPr phoneticPr
        {
            get
            {
                return this.phoneticPrField;
            }
            set
            {
                this.phoneticPrField = value;
            }
        }
        [XmlElement]
        public List<CT_ConditionalFormatting> conditionalFormatting
        {
            get
            {
                return this.conditionalFormattingField;
            }
            set
            {
                this.conditionalFormattingField = value;
            }
        }
        [XmlElement]
        public CT_DataValidations dataValidations
        {
            get
            {
                return this.dataValidationsField;
            }
            set
            {
                this.dataValidationsField = value;
            }
        }

        //[XmlArray(Order = 18)]
        [XmlElement("hyperlinks", IsNullable = false)]
        public CT_Hyperlinks hyperlinks
        {
            get
            {
                return this.hyperlinksField;
            }
            set
            {
                this.hyperlinksField = value;
            }
        }
        [XmlElement]
        public CT_PrintOptions printOptions
        {
            get
            {
                return this.printOptionsField;
            }
            set
            {
                this.printOptionsField = value;
            }
        }

        [XmlElement]
        public CT_PageMargins pageMargins
        {
            get
            {
                return this.pageMarginsField;
            }
            set
            {
                this.pageMarginsField = value;
            }
        }
        [XmlElement]
        public CT_PageSetup pageSetup
        {
            get
            {
                return this.pageSetupField;
            }
            set
            {
                this.pageSetupField = value;
            }
        }
        [XmlElement]
        public CT_HeaderFooter headerFooter
        {
            get
            {
                return this.headerFooterField;
            }
            set
            {
                this.headerFooterField = value;
            }
        }
        [XmlElement]
        public CT_PageBreak rowBreaks
        {
            get
            {
                return this.rowBreaksField;
            }
            set
            {
                this.rowBreaksField = value;
            }
        }
        [XmlElement]
        public CT_PageBreak colBreaks
        {
            get
            {
                return this.colBreaksField;
            }
            set
            {
                this.colBreaksField = value;
            }
        }

        //[XmlArray(Order = 25)]
        //[XmlArrayItem("customPr", IsNullable = false)]
        [XmlElement("customPr")]
        public List<CT_CustomProperty> customProperties
        {
            get
            {
                return this.customPropertiesField;
            }
            set
            {
                this.customPropertiesField = value;
            }
        }

        //[XmlArray(Order = 26)]
        //[XmlArrayItem("cellWatch", IsNullable = false)]
        [XmlElement("cellWatch")]
        public List<CT_CellWatch> cellWatches
        {
            get
            {
                return this.cellWatchesField;
            }
            set
            {
                this.cellWatchesField = value;
            }
        }
        [XmlElement]
        public CT_IgnoredErrors ignoredErrors
        {
            get
            {
                return this.ignoredErrorsField;
            }
            set
            {
                this.ignoredErrorsField = value;
            }
        }

        //[XmlArray(Order = 28)]
        //[XmlArrayItem("cellSmartTags", IsNullable = false)]
        [XmlElement("cellSmartTags")]
        public List<CT_CellSmartTags> smartTags
        {
            get
            {
                return this.smartTagsField;
            }
            set
            {
                this.smartTagsField = value;
            }
        }
        [XmlElement]
        public CT_Drawing drawing
        {
            get
            {
                return this.drawingField;
            }
            set
            {
                this.drawingField = value;
            }
        }
        [XmlElement]
        public CT_LegacyDrawing legacyDrawing
        {
            get
            {
                return this.legacyDrawingField;
            }
            set
            {
                this.legacyDrawingField = value;
            }
        }
        [XmlElement]
        public CT_LegacyDrawing legacyDrawingHF
        {
            get
            {
                return this.legacyDrawingHFField;
            }
            set
            {
                this.legacyDrawingHFField = value;
            }
        }
        [XmlElement]
        public CT_SheetBackgroundPicture picture
        {
            get
            {
                return this.pictureField;
            }
            set
            {
                this.pictureField = value;
            }
        }

        //[XmlArray(Order = 33)]
        //[XmlArrayItem("oleObject", IsNullable = false)]
        [XmlElement("oleObjects")]
        public List<CT_OleObject> oleObjects
        {
            get
            {
                return this.oleObjectsField;
            }
            set
            {
                this.oleObjectsField = value;
            }
        }

        //[XmlArray(Order = 34)]
        //[XmlArrayItem("control", IsNullable = false)]
        [XmlElement("control")]
        public List<CT_Control> controls
        {
            get
            {
                return this.controlsField;
            }
            set
            {
                this.controlsField = value;
            }
        }
        [XmlElement]
        public CT_WebPublishItems webPublishItems
        {
            get
            {
                return this.webPublishItemsField;
            }
            set
            {
                this.webPublishItemsField = value;
            }
        }
        [XmlElement]
        public CT_TableParts tableParts
        {
            get
            {
                return this.tablePartsField;
            }
            set
            {
                this.tablePartsField = value;
            }
        }
        [XmlElement]
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

        public void UnsetPageSetup()
        {
            throw new NotImplementedException();
        }
    }

}
