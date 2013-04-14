using NPOI.Util;

namespace NPOI.SS.Formula
{
    using System;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Util;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using System.Globalization;
    /**
     * Contains all the contextual information required to Evaluate an operation
     * within a formula
     *
     * For POI internal use only
     *
     * @author Josh Micich
     */
    public class OperationEvaluationContext
    {
        public static readonly FreeRefFunction UDF = UserDefinedFunction.instance;
        private IEvaluationWorkbook _workbook;
        private int _sheetIndex;
        private int _rowIndex;
        private int _columnIndex;
        private EvaluationTracker _tracker;
        private WorkbookEvaluator _bookEvaluator;

        public OperationEvaluationContext(WorkbookEvaluator bookEvaluator, IEvaluationWorkbook workbook, int sheetIndex, int srcRowNum,
                int srcColNum, EvaluationTracker tracker)
        {
            _bookEvaluator = bookEvaluator;
            _workbook = workbook;
            _sheetIndex = sheetIndex;
            _rowIndex = srcRowNum;
            _columnIndex = srcColNum;
            _tracker = tracker;
        }

        public IEvaluationWorkbook GetWorkbook()
        {
            return _workbook;
        }

        public int RowIndex
        {
            get
            {
                return _rowIndex;
            }
        }

        public int ColumnIndex
        {
            get
            {
                return _columnIndex;
            }
        }

        SheetRefEvaluator CreateExternSheetRefEvaluator(IExternSheetReferenceToken ptg)
        {
            return CreateExternSheetRefEvaluator(ptg.ExternSheetIndex);
        }
        SheetRefEvaluator CreateExternSheetRefEvaluator(int externSheetIndex)
        {
            ExternalSheet externalSheet = _workbook.GetExternalSheet(externSheetIndex);
            WorkbookEvaluator targetEvaluator;
            int otherSheetIndex;
            if (externalSheet == null)
            {
                // sheet is in same workbook
                otherSheetIndex = _workbook.ConvertFromExternSheetIndex(externSheetIndex);
                targetEvaluator = _bookEvaluator;
            }
            else
            {
                // look up sheet by name from external workbook
                String workbookName = externalSheet.GetWorkbookName();
                try
                {
                    targetEvaluator = _bookEvaluator.GetOtherWorkbookEvaluator(workbookName);
                }
                catch (WorkbookNotFoundException e)
                {
                    throw new RuntimeException(e.Message, e);
                }
                otherSheetIndex = targetEvaluator.GetSheetIndex(externalSheet.GetSheetName());
                if (otherSheetIndex < 0)
                {
                    throw new Exception("Invalid sheet name '" + externalSheet.GetSheetName()
                            + "' in bool '" + workbookName + "'.");
                }
            }
            return new SheetRefEvaluator(targetEvaluator, _tracker, otherSheetIndex);
        }

        /**
         * @return <code>null</code> if either workbook or sheet is not found
         */
        private SheetRefEvaluator CreateExternSheetRefEvaluator(String workbookName, String sheetName)
        {
            WorkbookEvaluator targetEvaluator;
            if (workbookName == null)
            {
                targetEvaluator = _bookEvaluator;
            }
            else
            {
                if (sheetName == null)
                {
                    throw new ArgumentException("sheetName must not be null if workbookName is provided");
                }
                try
                {
                    targetEvaluator = _bookEvaluator.GetOtherWorkbookEvaluator(workbookName);
                }
                catch (WorkbookNotFoundException)
                {
                    return null;
                }
            }
            int otherSheetIndex = sheetName == null ? _sheetIndex : targetEvaluator.GetSheetIndex(sheetName);
            if (otherSheetIndex < 0)
            {
                return null;
            }
            return new SheetRefEvaluator(targetEvaluator, _tracker, otherSheetIndex);
        }

        public SheetRefEvaluator GetRefEvaluatorForCurrentSheet()
        {
            return new SheetRefEvaluator(_bookEvaluator, _tracker, _sheetIndex);
        }



        /**
         * Resolves a cell or area reference dynamically.
         * @param workbookName the name of the workbook Containing the reference.  If <code>null</code>
         * the current workbook is assumed.  Note - to Evaluate formulas which use multiple workbooks,
         * a {@link CollaboratingWorkbooksEnvironment} must be set up.
         * @param sheetName the name of the sheet Containing the reference.  May be <code>null</code>
         * (when <c>workbookName</c> is also null) in which case the current workbook and sheet is
         * assumed.
         * @param refStrPart1 the single cell reference or first part of the area reference.  Must not
         * be <code>null</code>.
         * @param refStrPart2 the second part of the area reference. For single cell references this
         * parameter must be <code>null</code>
         * @param isA1Style specifies the format for <c>refStrPart1</c> and <c>refStrPart2</c>.
         * Pass <c>true</c> for 'A1' style and <c>false</c> for 'R1C1' style.
         * TODO - currently POI only supports 'A1' reference style
         * @return a {@link RefEval} or {@link AreaEval}
         */
        public ValueEval GetDynamicReference(String workbookName, String sheetName, String refStrPart1,
                String refStrPart2, bool isA1Style)
        {
            if (!isA1Style)
            {
                throw new Exception("R1C1 style not supported yet");
            }
            SheetRefEvaluator sre = CreateExternSheetRefEvaluator(workbookName, sheetName);
            if (sre == null)
            {
                return ErrorEval.REF_INVALID;
            }
            // ugly typecast - TODO - make spReadsheet version more easily accessible
            SpreadsheetVersion ssVersion = ((IFormulaParsingWorkbook)_workbook).GetSpreadsheetVersion();

            NameType part1refType = ClassifyCellReference(refStrPart1, ssVersion);
            switch (part1refType)
            {
                case NameType.BadCellOrNamedRange:
                    return ErrorEval.REF_INVALID;
                case NameType.NamedRange:
                    IEvaluationName nm = ((IFormulaParsingWorkbook)_workbook).GetName(refStrPart1, _sheetIndex);
                    if (!nm.IsRange)
                    {
                        throw new Exception("Specified name '" + refStrPart1 + "' is not a range as expected.");
                    }
                    return _bookEvaluator.EvaluateNameFormula(nm.NameDefinition, this);
            }
            if (refStrPart2 == null)
            {
                // no ':'
                switch (part1refType)
                {
                    case NameType.Column:
                    case NameType.Row:
                        return ErrorEval.REF_INVALID;
                    case NameType.Cell:
                        CellReference cr = new CellReference(refStrPart1);
                        return new LazyRefEval(cr.Row, cr.Col, sre);
                }
                throw new InvalidOperationException("Unexpected reference classification of '" + refStrPart1 + "'.");
            }
            NameType part2refType = ClassifyCellReference(refStrPart1, ssVersion);
            switch (part2refType)
            {
                case NameType.BadCellOrNamedRange:
                    return ErrorEval.REF_INVALID;
                case NameType.NamedRange:
                    throw new Exception("Cannot Evaluate '" + refStrPart1
                            + "'. Indirect Evaluation of defined names not supported yet");
            }

            if (part2refType != part1refType)
            {
                // LHS and RHS of ':' must be compatible
                return ErrorEval.REF_INVALID;
            }
            int firstRow, firstCol, lastRow, lastCol;
            switch (part1refType)
            {
                case NameType.Column:
                    firstRow = 0;
                    lastRow = ssVersion.LastRowIndex;
                    firstCol = ParseColRef(refStrPart1);
                    lastCol = ParseColRef(refStrPart2);
                    break;
                case NameType.Row:
                    firstCol = 0;
                    lastCol = ssVersion.LastColumnIndex;
                    firstRow = ParseRowRef(refStrPart1);
                    lastRow = ParseRowRef(refStrPart2);
                    break;
                case NameType.Cell:
                    CellReference cr;
                    cr = new CellReference(refStrPart1);
                    firstRow = cr.Row;
                    firstCol = cr.Col;
                    cr = new CellReference(refStrPart2);
                    lastRow = cr.Row;
                    lastCol = cr.Col;
                    break;
                default:
                    throw new InvalidOperationException("Unexpected reference classification of '" + refStrPart1 + "'.");
            }
            return new LazyAreaEval(firstRow, firstCol, lastRow, lastCol, sre);
        }

        private static int ParseRowRef(String refStrPart)
        {
            return CellReference.ConvertColStringToIndex(refStrPart);
        }

        private static int ParseColRef(String refStrPart)
        {
            return Int32.Parse(refStrPart, CultureInfo.InvariantCulture) - 1;
        }

        private static NameType ClassifyCellReference(String str, SpreadsheetVersion ssVersion)
        {
            int len = str.Length;
            if (len < 1)
            {
                return NameType.BadCellOrNamedRange;
            }
            return CellReference.ClassifyCellReference(str, ssVersion);
        }

        public FreeRefFunction FindUserDefinedFunction(String functionName)
        {
            return _bookEvaluator.FindUserDefinedFunction(functionName);
        }

        public ValueEval GetRefEval(int rowIndex, int columnIndex)
        {
            SheetRefEvaluator sre = GetRefEvaluatorForCurrentSheet();
            return new LazyRefEval(rowIndex, columnIndex, sre);
        }
        public ValueEval GetRef3DEval(int rowIndex, int columnIndex, int extSheetIndex)
        {
            SheetRefEvaluator sre = CreateExternSheetRefEvaluator(extSheetIndex);
            return new LazyRefEval(rowIndex, columnIndex, sre);
        }
        public ValueEval GetAreaEval(int firstRowIndex, int firstColumnIndex,
                int lastRowIndex, int lastColumnIndex)
        {
            SheetRefEvaluator sre = GetRefEvaluatorForCurrentSheet();
            return new LazyAreaEval(firstRowIndex, firstColumnIndex, lastRowIndex, lastColumnIndex, sre);
        }
        public ValueEval GetArea3DEval(int firstRowIndex, int firstColumnIndex,
                int lastRowIndex, int lastColumnIndex, int extSheetIndex)
        {
            SheetRefEvaluator sre = CreateExternSheetRefEvaluator(extSheetIndex);
            return new LazyAreaEval(firstRowIndex, firstColumnIndex, lastRowIndex, lastColumnIndex, sre);
        }
        public ValueEval GetNameXEval(NameXPtg nameXPtg)
        {
            ExternalSheet externSheet = _workbook.GetExternalSheet(nameXPtg.SheetRefIndex);
            if (externSheet == null)
                return new NameXEval(nameXPtg);
            String workbookName = externSheet.GetWorkbookName();
            ExternalName externName = _workbook.GetExternalName(
                  nameXPtg.SheetRefIndex,
                  nameXPtg.NameIndex
            );
            try
            {
                WorkbookEvaluator refWorkbookEvaluator = _bookEvaluator.GetOtherWorkbookEvaluator(workbookName);
                IEvaluationName evaluationName = refWorkbookEvaluator.GetName(externName.Name, externName.Ix - 1);
                if (evaluationName != null && evaluationName.HasFormula)
                {
                    if (evaluationName.NameDefinition.Length > 1)
                    {
                        throw new Exception("Complex name formulas not supported yet");
                    }
                    Ptg ptg = evaluationName.NameDefinition[0];
                    if (ptg is Ref3DPtg)
                    {
                        Ref3DPtg ref3D = (Ref3DPtg)ptg;
                        int sheetIndex = refWorkbookEvaluator.GetSheetIndexByExternIndex(ref3D.ExternSheetIndex);
                        String sheetName = refWorkbookEvaluator.GetSheetName(sheetIndex);
                        SheetRefEvaluator sre = CreateExternSheetRefEvaluator(workbookName, sheetName);
                        return new LazyRefEval(ref3D.Row, ref3D.Column, sre);
                    }
                    else if (ptg is Area3DPtg)
                    {
                        Area3DPtg area3D = (Area3DPtg)ptg;
                        int sheetIndex = refWorkbookEvaluator.GetSheetIndexByExternIndex(area3D.ExternSheetIndex);
                        String sheetName = refWorkbookEvaluator.GetSheetName(sheetIndex);
                        SheetRefEvaluator sre = CreateExternSheetRefEvaluator(workbookName, sheetName);
                        return new LazyAreaEval(area3D.FirstRow, area3D.FirstColumn, area3D.LastRow, area3D.LastColumn, sre);
                    }
                }
                return ErrorEval.REF_INVALID;
            }
            catch (WorkbookNotFoundException)
            {
                return ErrorEval.REF_INVALID;
            }
        }
    }
}