using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
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

        private List<CT_Cols> colsField = new List<CT_Cols>();

        private CT_SheetData sheetDataField = new CT_SheetData();

        private CT_SheetCalcPr sheetCalcPrField = null;

        private CT_SheetProtection sheetProtectionField = null;

        private CT_ProtectedRanges protectedRangesField = null;

        private CT_Scenarios scenariosField = null;

        private CT_AutoFilter autoFilterField = null;

        private CT_SortState sortStateField = null;

        private CT_DataConsolidate dataConsolidateField = null;

        private CT_CustomSheetViews customSheetViewsField = null;

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

        private CT_CustomProperties customPropertiesField = null;

        private CT_CellWatches cellWatchesField = null;

        private CT_IgnoredErrors ignoredErrorsField = null;

        private CT_CellSmartTags smartTagsField = null;

        private CT_Drawing drawingField = null;

        private CT_LegacyDrawing legacyDrawingField = null;

        private CT_LegacyDrawing legacyDrawingHFField = null;

        private CT_SheetBackgroundPicture pictureField = null;

        private CT_OleObjects oleObjectsField = null;

        private CT_Controls controlsField = null;

        private CT_WebPublishItems webPublishItemsField = null;

        private CT_TableParts tablePartsField = null;

        private CT_ExtensionList extLstField = null;

        public static CT_Worksheet Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Worksheet ctObj = new CT_Worksheet();
            ctObj.cols = new List<CT_Cols>();
            ctObj.conditionalFormatting = new List<CT_ConditionalFormatting>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "sheetPr")
                    ctObj.sheetPr = CT_SheetPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dimension")
                    ctObj.dimension = CT_SheetDimension.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sheetViews")
                    ctObj.sheetViews = CT_SheetViews.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sheetFormatPr")
                    ctObj.sheetFormatPr = CT_SheetFormatPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sheetData")
                    ctObj.sheetData = CT_SheetData.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sheetCalcPr")
                    ctObj.sheetCalcPr = CT_SheetCalcPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sheetProtection")
                    ctObj.sheetProtection = CT_SheetProtection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "protectedRanges")
                    ctObj.protectedRanges = CT_ProtectedRanges.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "scenarios")
                    ctObj.scenarios = CT_Scenarios.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "autoFilter")
                    ctObj.autoFilter = CT_AutoFilter.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sortState")
                    ctObj.sortState = CT_SortState.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dataConsolidate")
                    ctObj.dataConsolidate = CT_DataConsolidate.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "customSheetViews")
                    ctObj.customSheetViews = CT_CustomSheetViews.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "mergeCells")
                    ctObj.mergeCells = CT_MergeCells.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "phoneticPr")
                    ctObj.phoneticPr = CT_PhoneticPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dataValidations")
                    ctObj.dataValidations = CT_DataValidations.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hyperlinks")
                    ctObj.hyperlinks = CT_Hyperlinks.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "printOptions")
                    ctObj.printOptions = CT_PrintOptions.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pageMargins")
                    ctObj.pageMargins = CT_PageMargins.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pageSetup")
                    ctObj.pageSetup = CT_PageSetup.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "headerFooter")
                    ctObj.headerFooter = CT_HeaderFooter.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rowBreaks")
                    ctObj.rowBreaks = CT_PageBreak.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "colBreaks")
                    ctObj.colBreaks = CT_PageBreak.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "customProperties")
                    ctObj.customProperties = CT_CustomProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cellWatches")
                    ctObj.cellWatches = CT_CellWatches.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "ignoredErrors")
                    ctObj.ignoredErrors = CT_IgnoredErrors.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "smartTags")
                    ctObj.smartTags = CT_CellSmartTags.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "drawing")
                    ctObj.drawing = CT_Drawing.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "legacyDrawing")
                    ctObj.legacyDrawing = CT_LegacyDrawing.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "legacyDrawingHF")
                    ctObj.legacyDrawingHF = CT_LegacyDrawing.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "picture")
                    ctObj.picture = CT_SheetBackgroundPicture.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "oleObjects")
                    ctObj.oleObjects = CT_OleObjects.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "controls")
                    ctObj.controls = CT_Controls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "webPublishItems")
                    ctObj.webPublishItems = CT_WebPublishItems.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tableParts")
                    ctObj.tableParts = CT_TableParts.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cols")
                    ctObj.cols.Add(CT_Cols.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "conditionalFormatting")
                    ctObj.conditionalFormatting.Add(CT_ConditionalFormatting.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sw.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                if (this.sheetPr != null)
                    this.sheetPr.Write(sw, "sheetPr");
                if (this.dimension != null)
                    this.dimension.Write(sw, "dimension");
                if (this.sheetViews != null)
                    this.sheetViews.Write(sw, "sheetViews");
                if (this.sheetFormatPr != null)
                    this.sheetFormatPr.Write(sw, "sheetFormatPr");
                if (this.cols != null)
                {
                    foreach (CT_Cols x in this.cols)
                    {
                        x.Write(sw, "cols");
                    }
                }
                if (this.sheetData != null)
                    this.sheetData.Write(sw, "sheetData");
                if (this.sheetCalcPr != null)
                    this.sheetCalcPr.Write(sw, "sheetCalcPr");
                if (this.sheetProtection != null)
                    this.sheetProtection.Write(sw, "sheetProtection");
                if (this.protectedRanges != null)
                    this.protectedRanges.Write(sw, "protectedRanges");
                if (this.scenarios != null)
                    this.scenarios.Write(sw, "scenarios");
                if (this.autoFilter != null)
                    this.autoFilter.Write(sw, "autoFilter");
                if (this.sortState != null)
                    this.sortState.Write(sw, "sortState");
                if (this.dataConsolidate != null)
                    this.dataConsolidate.Write(sw, "dataConsolidate");
                if (this.customSheetViews != null)
                    this.customSheetViews.Write(sw, "customSheetViews");
                if (this.mergeCells != null)
                    this.mergeCells.Write(sw, "mergeCells");
                if (this.phoneticPr != null)
                    this.phoneticPr.Write(sw, "phoneticPr");
                if (this.conditionalFormatting != null)
                {
                    foreach (CT_ConditionalFormatting x in this.conditionalFormatting)
                    {
                        x.Write(sw, "conditionalFormatting");
                    }
                }
                if (this.dataValidations != null)
                    this.dataValidations.Write(sw, "dataValidations");
                if (this.hyperlinks != null)
                    this.hyperlinks.Write(sw, "hyperlinks");
                if (this.printOptions != null)
                    this.printOptions.Write(sw, "printOptions");
                if (this.pageMargins != null)
                    this.pageMargins.Write(sw, "pageMargins");
                if (this.pageSetup != null)
                    this.pageSetup.Write(sw, "pageSetup");
                if (this.headerFooter != null)
                    this.headerFooter.Write(sw, "headerFooter");
                if (this.rowBreaks != null)
                    this.rowBreaks.Write(sw, "rowBreaks");
                if (this.colBreaks != null)
                    this.colBreaks.Write(sw, "colBreaks");
                if (this.customProperties != null)
                    this.customProperties.Write(sw, "customProperties");
                if (this.cellWatches != null)
                    this.cellWatches.Write(sw, "cellWatches");
                if (this.ignoredErrors != null)
                    this.ignoredErrors.Write(sw, "ignoredErrors");
                if (this.smartTags != null)
                    this.smartTags.Write(sw, "smartTags");
                if (this.drawing != null)
                    this.drawing.Write(sw, "drawing");
                if (this.legacyDrawing != null)
                    this.legacyDrawing.Write(sw, "legacyDrawing");
                if (this.legacyDrawingHF != null)
                    this.legacyDrawingHF.Write(sw, "legacyDrawingHF");
                if (this.picture != null)
                    this.picture.Write(sw, "picture");
                if (this.oleObjects != null)
                    this.oleObjects.Write(sw, "oleObjects");
                if (this.controls != null)
                    this.controls.Write(sw, "controls");
                if (this.webPublishItems != null)
                    this.webPublishItems.Write(sw, "webPublishItems");
                if (this.tableParts != null)
                    this.tableParts.Write(sw, "tableParts");
                if (this.extLst != null)
                    this.extLst.Write(sw, "extLst");
                sw.Write("</worksheet>");
            }
        }

        public CT_AutoFilter AddNewAutoFilter()
        {
            this.autoFilterField = new CT_AutoFilter();
            return this.autoFilterField;
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
            if (null == colsField)
            {
                colsField = new List<CT_Cols>();
                colsField.Add(new CT_Cols());
            }
            return this.colsField[index];
        }
        public List<CT_Cols> GetColsList()
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

        [XmlElement]
        public CT_ProtectedRanges protectedRanges
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

        [XmlElement]
        public CT_CustomSheetViews customSheetViews
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
                if (this.conditionalFormattingField == null)
                    this.conditionalFormattingField = new List<CT_ConditionalFormatting>();
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
        [XmlElement]
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

        [XmlElement]
        public CT_CustomProperties customProperties
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

        [XmlElement]
        public CT_CellWatches cellWatches
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
        [XmlElement]
        public CT_CellSmartTags smartTags
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

        [XmlElement]
        public CT_OleObjects oleObjects
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

        [XmlElement]
        public CT_Controls controls
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
            this.pageSetup = null;
        }
    }

}
