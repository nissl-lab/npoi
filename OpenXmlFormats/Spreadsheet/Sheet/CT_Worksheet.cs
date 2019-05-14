using System;
using System.Linq;
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
        private List<CT_ConditionalFormatting> conditionalFormattingField = null;

        public static CT_Worksheet Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Worksheet ctObj = new CT_Worksheet();
            ctObj.cols = new List<CT_Cols>();
            ctObj.conditionalFormatting = new List<CT_ConditionalFormatting>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(sheetPr))
                    ctObj.sheetPr = CT_SheetPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(dimension))
                    ctObj.dimension = CT_SheetDimension.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(sheetViews))
                    ctObj.sheetViews = CT_SheetViews.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(sheetFormatPr))
                    ctObj.sheetFormatPr = CT_SheetFormatPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(sheetData))
                    ctObj.sheetData = CT_SheetData.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(sheetCalcPr))
                    ctObj.sheetCalcPr = CT_SheetCalcPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(sheetProtection))
                    ctObj.sheetProtection = CT_SheetProtection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(protectedRanges))
                    ctObj.protectedRanges = CT_ProtectedRanges.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(scenarios))
                    ctObj.scenarios = CT_Scenarios.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(autoFilter))
                    ctObj.autoFilter = CT_AutoFilter.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(sortState))
                    ctObj.sortState = CT_SortState.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(dataConsolidate))
                    ctObj.dataConsolidate = CT_DataConsolidate.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(customSheetViews))
                    ctObj.customSheetViews = CT_CustomSheetViews.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(mergeCells))
                    ctObj.mergeCells = CT_MergeCells.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(phoneticPr))
                    ctObj.phoneticPr = CT_PhoneticPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(dataValidations))
                    ctObj.dataValidations = CT_DataValidations.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(hyperlinks))
                    ctObj.hyperlinks = CT_Hyperlinks.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(printOptions))
                    ctObj.printOptions = CT_PrintOptions.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(pageMargins))
                    ctObj.pageMargins = CT_PageMargins.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(pageSetup))
                    ctObj.pageSetup = CT_PageSetup.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(headerFooter))
                    ctObj.headerFooter = CT_HeaderFooter.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(rowBreaks))
                    ctObj.rowBreaks = CT_PageBreak.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(colBreaks))
                    ctObj.colBreaks = CT_PageBreak.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(customProperties))
                    ctObj.customProperties = CT_CustomProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(cellWatches))
                    ctObj.cellWatches = CT_CellWatches.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(ignoredErrors))
                    ctObj.ignoredErrors = CT_IgnoredErrors.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(smartTags))
                    ctObj.smartTags = CT_CellSmartTags.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(drawing))
                    ctObj.drawing = CT_Drawing.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(legacyDrawing))
                    ctObj.legacyDrawing = CT_LegacyDrawing.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(legacyDrawingHF))
                    ctObj.legacyDrawingHF = CT_LegacyDrawing.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(picture))
                    ctObj.picture = CT_SheetBackgroundPicture.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(oleObjects))
                    ctObj.oleObjects = CT_OleObjects.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(controls))
                    ctObj.controls = CT_Controls.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(webPublishItems))
                    ctObj.webPublishItems = CT_WebPublishItems.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(tableParts))
                    ctObj.tableParts = CT_TableParts.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(cols))
                    ctObj.cols.Add(CT_Cols.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(conditionalFormatting))
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
                this.sheetPr?.Write(sw, nameof(sheetPr));
                this.dimension?.Write(sw, nameof(dimension));
                this.sheetViews?.Write(sw, nameof(sheetViews));
                this.sheetFormatPr?.Write(sw, nameof(sheetFormatPr));
                this.cols?.ForEach(x => x.Write(sw, nameof(cols)));
                this.sheetData?.Write(sw, nameof(sheetData));
                this.sheetCalcPr?.Write(sw, nameof(sheetCalcPr));
                this.sheetProtection?.Write(sw, nameof(sheetProtection));
                this.protectedRanges?.Write(sw, nameof(protectedRanges));
                this.scenarios?.Write(sw, nameof(scenarios));
                this.autoFilter?.Write(sw, nameof(autoFilter));
                this.sortState?.Write(sw, nameof(sortState));
                this.dataConsolidate?.Write(sw, nameof(dataConsolidate));
                this.customSheetViews?.Write(sw, nameof(customSheetViews));
                this.mergeCells?.Write(sw, nameof(mergeCells));
                this.phoneticPr?.Write(sw, nameof(phoneticPr));
                this.conditionalFormatting?.ForEach(x => x.Write(sw, nameof(conditionalFormatting)));
                this.dataValidations?.Write(sw, nameof(dataValidations));
                this.hyperlinks?.Write(sw, nameof(hyperlinks));
                this.printOptions?.Write(sw, nameof(printOptions));
                this.pageMargins?.Write(sw, nameof(pageMargins));
                this.pageSetup?.Write(sw, nameof(pageSetup));
                this.headerFooter?.Write(sw, nameof(headerFooter));
                this.rowBreaks?.Write(sw, nameof(rowBreaks));
                this.colBreaks?.Write(sw, nameof(colBreaks));
                this.customProperties?.Write(sw, nameof(customProperties));
                this.cellWatches?.Write(sw, nameof(cellWatches));
                this.ignoredErrors?.Write(sw, nameof(ignoredErrors));
                this.smartTags?.Write(sw, nameof(smartTags));
                this.drawing?.Write(sw, nameof(drawing));
                this.legacyDrawing?.Write(sw, nameof(legacyDrawing));
                this.legacyDrawingHF?.Write(sw, nameof(legacyDrawingHF));
                this.picture?.Write(sw, nameof(picture));
                this.oleObjects?.Write(sw, nameof(oleObjects));
                this.controls?.Write(sw, nameof(controls));
                this.webPublishItems?.Write(sw, nameof(webPublishItems));
                this.tableParts?.Write(sw, nameof(tableParts));
                this.extLst?.Write(sw, nameof(extLst));
                sw.Write("</worksheet>");
            }
        }

        public CT_AutoFilter AddNewAutoFilter()
        {
            this.autoFilter = new CT_AutoFilter();
            return this.autoFilter;
        }
        public bool IsSetRowBreaks()
        {
            return this.rowBreaks != null;
        }
        public CT_Drawing AddNewDrawing()
        {
            this.drawing = new CT_Drawing();
            return drawing;
        }
        public CT_LegacyDrawing AddNewLegacyDrawing()
        {
            this.legacyDrawing = new CT_LegacyDrawing();
            return legacyDrawing;
        }
        public CT_PageBreak AddNewRowBreaks()
        {
            this.rowBreaks = new CT_PageBreak();
            return this.rowBreaks;
        }
        public CT_PageBreak AddNewColBreaks()
        {
            this.colBreaks = new CT_PageBreak();
            return this.colBreaks;
        }
        public bool IsSetSheetFormatPr()
        {
            return this.sheetFormatPr != null;
        }
        public bool IsSetPrintOptions()
        {
            return this.printOptions != null;
        }
        public void UnsetMergeCells()
        {
            this.mergeCells = null;
        }
        public CT_PrintOptions AddNewPrintOptions()
        {
            this.printOptions = new CT_PrintOptions();
            return this.printOptions;
        }
        public CT_DataValidations AddNewDataValidations()
        {
            this.dataValidations = new CT_DataValidations();
            return this.dataValidations;
        }
        public CT_SheetViews AddNewSheetViews()
        {
            this.sheetViews = new CT_SheetViews();
            return this.sheetViews;
        }
        public CT_Hyperlinks AddNewHyperlinks()
        {
            this.hyperlinks = new CT_Hyperlinks();
            return this.hyperlinks;
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
            return this.conditionalFormattingField?[index];
        }
        public CT_MergeCells AddNewMergeCells()
        {
            this.mergeCells = new CT_MergeCells();
            return this.mergeCells;
        }
        public bool IsSetColBreaks()
        {
            return this.colBreaks != null;
        }
        public bool IsSetHyperlinks()
        {
            return this.hyperlinks != null;
        }
        public bool IsSetMergeCells()
        {
            return this.mergeCells != null;
        }
        public bool IsSetSheetProtection()
        {
            return this.sheetProtection != null;
        }
        public bool IsSetDrawing()
        {
            return this.drawing != null;
        }
        public void UnsetDrawing()
        {
            this.drawing = null;
        }
        public bool IsSetLegacyDrawing()
        {
            return this.legacyDrawing != null;
        }
        public void UnsetLegacyDrawing()
        {
            this.legacyDrawing = null;
        }
        public bool IsSetPageSetup()
        {
            return this.pageSetup != null;
        }
        public bool IsSetTableParts()
        {
            return this.tableParts != null;
        }
        public bool IsSetSheetCalcPr()
        {
            return this.sheetCalcPr != null;
        }
        public CT_SheetProtection AddNewSheetProtection()
        {
            this.sheetProtection = new CT_SheetProtection();
            return this.sheetProtection;
        }
        public CT_TableParts AddNewTableParts()
        {
            this.tableParts = new CT_TableParts();
            return this.tableParts;
        }
        public CT_PageMargins AddNewPageMargins()
        {
            this.pageMargins = new CT_PageMargins();
            return this.pageMargins;
        }
        public CT_PageSetup AddNewPageSetup()
        {
            this.pageSetup = new CT_PageSetup();
            return this.pageSetup;
        }
        public void SetColsArray(List<CT_Cols> a)
        {
            this.cols = a;
        }
        public int sizeOfColsArray()
        {
            return (null == cols) ? 0 : this.cols.Count;
        }
        public void RemoveCols(int index)
        {
            this.cols.RemoveAt(index);
        }
        public CT_Cols AddNewCols()
        {
            if (null == cols) { cols = new List<CT_Cols>(); }
            CT_Cols newCols = new CT_Cols();
            this.cols.Add(newCols);
            return newCols;
        }
        public void SetColsArray(int index, CT_Cols newCols)
        {
            if (null == cols)
            {
                cols = new List<CT_Cols>();
            }
            else
            {
                cols.Clear();
            }
            this.cols.Insert(index, newCols);
        }
        public CT_Cols GetColsArray(int index)
        {
            if (null == cols)
            {
                cols = new List<CT_Cols>();
                cols.Add(new CT_Cols());
            }
            return this.cols[index];
        }
        public List<CT_Cols> GetColsList()
        {
            return this.cols;
        }
        public bool IsSetPageMargins()
        {
            return this.pageMargins != null;
        }
        public bool IsSetHyperLinks()
        {
            return this.hyperlinks != null;
        }
        public bool IsSetSheetPr()
        {
            return this.sheetPr != null;
        }
        public int SizeOfConditionalFormattingArray()
        {
            return this.conditionalFormattingField?.Count ?? 0;
            return (null == conditionalFormattingField) ? 0 : this.conditionalFormatting.Count;
        }

        public void UnsetSheetProtection()
        {
            this.sheetProtection = null;
        }

        public CT_SheetFormatPr AddNewSheetFormatPr()
        {
            this.sheetFormatPr = new CT_SheetFormatPr();
            return sheetFormatPr;
        }
        public CT_SheetCalcPr AddNewSheetCalcPr()
        {
            this.sheetCalcPr = new CT_SheetCalcPr();
            return sheetCalcPr;
        }
        public CT_SheetPr AddNewSheetPr()
        {
            this.sheetPr = new CT_SheetPr();
            return sheetPr;
        }
        public CT_SheetDimension AddNewDimension()
        {
            this.dimension = new CT_SheetDimension();
            return dimension;
        }
        public CT_SheetData AddNewSheetData()
        {
            this.sheetData = new CT_SheetData();
            return sheetData;
        }
        [XmlElement(nameof(sheetPr))]//, Order=0)]
        public CT_SheetPr sheetPr { get; set; } = null;

        [XmlElement]
        public CT_SheetDimension dimension { get; set; } = null;
        [XmlElement]
        public CT_SheetViews sheetViews { get; set; } = null;
        [XmlElement]
        public CT_SheetFormatPr sheetFormatPr { get; set; } = null;

        //[XmlArray(Order = 4)]
        // [XmlArrayItem("cols", typeof(CT_Cols), IsNullable = false)]
        [XmlElement(nameof(cols))]
        public List<CT_Cols> cols { get; set; } = new List<CT_Cols>();

        [XmlElement(nameof(sheetData), IsNullable = false)]
        public CT_SheetData sheetData { get; set; } = new CT_SheetData();
        [XmlElement]
        public CT_SheetCalcPr sheetCalcPr { get; set; } = null;
        [XmlElement]
        public CT_SheetProtection sheetProtection { get; set; } = null;

        [XmlElement]
        public CT_ProtectedRanges protectedRanges { get; set; } = null;

        public CT_Scenarios scenarios { get; set; } = null;

        public CT_AutoFilter autoFilter { get; set; } = null;

        public CT_SortState sortState { get; set; } = null;

        public CT_DataConsolidate dataConsolidate { get; set; } = null;

        [XmlElement]
        public CT_CustomSheetViews customSheetViews { get; set; } = null;
        [XmlElement]
        public CT_MergeCells mergeCells { get; set; } = null;
        [XmlElement]
        public CT_PhoneticPr phoneticPr { get; set; } = null;
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
        public CT_DataValidations dataValidations { get; set; } = null;

        //[XmlArray(Order = 18)]
        [XmlElement]
        public CT_Hyperlinks hyperlinks { get; set; } = null;
        [XmlElement]
        public CT_PrintOptions printOptions { get; set; } = null;

        [XmlElement]
        public CT_PageMargins pageMargins { get; set; } = null;
        [XmlElement]
        public CT_PageSetup pageSetup { get; set; } = null;
        [XmlElement]
        public CT_HeaderFooter headerFooter { get; set; } = null;
        [XmlElement]
        public CT_PageBreak rowBreaks { get; set; } = null;
        [XmlElement]
        public CT_PageBreak colBreaks { get; set; } = null;

        [XmlElement]
        public CT_CustomProperties customProperties { get; set; } = null;

        [XmlElement]
        public CT_CellWatches cellWatches { get; set; } = null;
        [XmlElement]
        public CT_IgnoredErrors ignoredErrors { get; set; } = null;
        [XmlElement]
        public CT_CellSmartTags smartTags { get; set; } = null;
        [XmlElement]
        public CT_Drawing drawing { get; set; } = null;
        [XmlElement]
        public CT_LegacyDrawing legacyDrawing { get; set; } = null;
        [XmlElement]
        public CT_LegacyDrawing legacyDrawingHF { get; set; } = null;
        [XmlElement]
        public CT_SheetBackgroundPicture picture { get; set; } = null;

        [XmlElement]
        public CT_OleObjects oleObjects { get; set; } = null;

        [XmlElement]
        public CT_Controls controls { get; set; } = null;
        [XmlElement]
        public CT_WebPublishItems webPublishItems { get; set; } = null;
        [XmlElement]
        public CT_TableParts tableParts { get; set; } = null;
        [XmlElement]
        public CT_ExtensionList extLst { get; set; } = null;

        public void UnsetPageSetup()
        {
            this.pageSetup = null;
        }
    }

}
