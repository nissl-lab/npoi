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

        SheetRangeEvaluator CreateExternSheetRefEvaluator(IExternSheetReferenceToken ptg)
        {
            return CreateExternSheetRefEvaluator(ptg.ExternSheetIndex);
        }

        SheetRangeEvaluator CreateExternSheetRefEvaluator(String firstSheetName, String lastSheetName, int externalWorkbookNumber)
        {
            ExternalSheet externalSheet = _workbook.GetExternalSheet(firstSheetName, lastSheetName, externalWorkbookNumber);
            return CreateExternSheetRefEvaluator(externalSheet);
        }

        SheetRangeEvaluator CreateExternSheetRefEvaluator(int externSheetIndex)
        {
            ExternalSheet externalSheet = _workbook.GetExternalSheet(externSheetIndex);
            return CreateExternSheetRefEvaluator(externalSheet);
        }
        SheetRangeEvaluator CreateExternSheetRefEvaluator(ExternalSheet externalSheet)
        {
            WorkbookEvaluator targetEvaluator;
            int otherFirstSheetIndex;
            int otherLastSheetIndex = -1;
            if (externalSheet == null || externalSheet.WorkbookName == null)
            {
                // sheet is in same workbook
                targetEvaluator = _bookEvaluator;
                otherFirstSheetIndex = _workbook.GetSheetIndex(externalSheet.SheetName);

                if (externalSheet is ExternalSheetRange)
                {
                    String lastSheetName = ((ExternalSheetRange)externalSheet).LastSheetName;
                    otherLastSheetIndex = _workbook.GetSheetIndex(lastSheetName);
                }
            }
            else
            {
                // look up sheet by name from external workbook
                String workbookName = externalSheet.WorkbookName;
                try
                {
                    targetEvaluator = _bookEvaluator.GetOtherWorkbookEvaluator(workbookName);
                }
                catch (WorkbookNotFoundException e)
                {
                    throw new RuntimeException(e.Message, e);
                }

                otherFirstSheetIndex = targetEvaluator.GetSheetIndex(externalSheet.SheetName);
                if (externalSheet is ExternalSheetRange)
                {
                    String lastSheetName = ((ExternalSheetRange)externalSheet).LastSheetName;
                    otherLastSheetIndex = targetEvaluator.GetSheetIndex(lastSheetName);
                }

                if (otherFirstSheetIndex < 0)
                {
                    throw new Exception("Invalid sheet name '" + externalSheet.SheetName
                            + "' in bool '" + workbookName + "'.");
                }
            }

            if (otherLastSheetIndex == -1)
            {
                // Reference to just one sheet
                otherLastSheetIndex = otherFirstSheetIndex;
            }

            SheetRefEvaluator[] Evals = new SheetRefEvaluator[otherLastSheetIndex - otherFirstSheetIndex + 1];
            for (int i = 0; i < Evals.Length; i++)
            {
                int otherSheetIndex = i + otherFirstSheetIndex;
                Evals[i] = new SheetRefEvaluator(targetEvaluator, _tracker, otherSheetIndex);
            }
            return new SheetRangeEvaluator(otherFirstSheetIndex, otherLastSheetIndex, Evals);
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

        public SheetRangeEvaluator GetRefEvaluatorForCurrentSheet()
        {
            SheetRefEvaluator sre = new SheetRefEvaluator(_bookEvaluator, _tracker, _sheetIndex);
            return new SheetRangeEvaluator(_sheetIndex, sre);
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
            SheetRefEvaluator se = CreateExternSheetRefEvaluator(workbookName, sheetName);
            if (se == null)
            {
                return ErrorEval.REF_INVALID;
            }
            SheetRangeEvaluator sre = new SheetRangeEvaluator(_sheetIndex, se);

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
                    if (part2refType.Equals(NameType.Column))
                    {
                        lastRow = ssVersion.LastRowIndex;
                        firstCol = ParseRowRef(refStrPart1);
                        lastCol = ParseRowRef(refStrPart2);
                    }
                    else
                    {
                        lastRow = ssVersion.LastRowIndex;
                        firstCol = ParseColRef(refStrPart1);
                        lastCol = ParseColRef(refStrPart2);
                    }
                    break;
                case NameType.Row:
                    // support of cell range in the form of integer:integer
                    firstCol = 0;
                    if (part2refType.Equals(NameType.Row))
                    {
                        firstRow = ParseColRef(refStrPart1);
                        lastRow = ParseColRef(refStrPart2);
                        lastCol = ssVersion.LastColumnIndex;
                    }
                    else
                    {
                        lastCol = ssVersion.LastColumnIndex;
                        firstRow = ParseRowRef(refStrPart1);
                        lastRow = ParseRowRef(refStrPart2);
                    }
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
            SheetRangeEvaluator sre = GetRefEvaluatorForCurrentSheet();
            return new LazyRefEval(rowIndex, columnIndex, sre);
        }
        public ValueEval GetRef3DEval(Ref3DPtg rptg)
        {
            SheetRangeEvaluator sre = CreateExternSheetRefEvaluator(rptg.ExternSheetIndex);
            return new LazyRefEval(rptg.Row, rptg.Column, sre);
        }
        public ValueEval GetRef3DEval(Ref3DPxg rptg)
        {
            SheetRangeEvaluator sre = CreateExternSheetRefEvaluator(rptg.SheetName, rptg.LastSheetName, rptg.ExternalWorkbookNumber);
            return new LazyRefEval(rptg.Row, rptg.Column, sre);
        }
        public ValueEval GetAreaEval(int firstRowIndex, int firstColumnIndex,
                int lastRowIndex, int lastColumnIndex)
        {
            SheetRangeEvaluator sre = GetRefEvaluatorForCurrentSheet();
            return new LazyAreaEval(firstRowIndex, firstColumnIndex, lastRowIndex, lastColumnIndex, sre);
        }
        public ValueEval GetArea3DEval(Area3DPtg aptg)
        {
            SheetRangeEvaluator sre = CreateExternSheetRefEvaluator(aptg.ExternSheetIndex);
            return new LazyAreaEval(aptg.FirstRow, aptg.FirstColumn,
                    aptg.LastRow, aptg.LastColumn, sre);
        }
        public ValueEval GetArea3DEval(Area3DPxg aptg)
        {
            SheetRangeEvaluator sre = CreateExternSheetRefEvaluator(aptg.SheetName, aptg.LastSheetName, aptg.ExternalWorkbookNumber);
            return new LazyAreaEval(aptg.FirstRow, aptg.FirstColumn,
                    aptg.LastRow, aptg.LastColumn, sre);
        }
        public ValueEval GetNameXEval(NameXPtg nameXPtg)
        {
            ExternalSheet externSheet = _workbook.GetExternalSheet(nameXPtg.SheetRefIndex);
            if (externSheet == null || externSheet.WorkbookName == null)
            {
                // External reference to our own workbook's name
                return GetLocalNameXEval(nameXPtg);
            }
            // Look it up for the external workbook
            String workbookName = externSheet.WorkbookName;
            ExternalName externName = _workbook.GetExternalName(
                  nameXPtg.SheetRefIndex,
                  nameXPtg.NameIndex
            );
            return GetExternalNameXEval(externName, workbookName);
        }
        public ValueEval GetNameXEval(NameXPxg nameXPxg)
        {
            ExternalSheet externSheet = _workbook.GetExternalSheet(nameXPxg.SheetName, null, nameXPxg.ExternalWorkbookNumber);
            if (externSheet == null || externSheet.WorkbookName == null)
            {
                // External reference to our own workbook's name
                return GetLocalNameXEval(nameXPxg);
            }

            // Look it up for the external workbook
            String workbookName = externSheet.WorkbookName;
            ExternalName externName = _workbook.GetExternalName(
                  nameXPxg.NameName,
                  nameXPxg.SheetName,
                  nameXPxg.ExternalWorkbookNumber
            );
            return GetExternalNameXEval(externName, workbookName);
        }
        private ValueEval GetLocalNameXEval(NameXPxg nameXPxg)
        {
            // Look up the sheet, if present
            int sIdx = -1;
            if (nameXPxg.SheetName != null)
            {
                sIdx = _workbook.GetSheetIndex(nameXPxg.SheetName);
            }

            // Is it a name or a function?
            String name = nameXPxg.NameName;
            IEvaluationName evalName = _workbook.GetName(name, sIdx);
            if (evalName != null)
            {
                // Process it as a name
                return new ExternalNameEval(evalName);
            }
            else
            {
                // Must be an external function
                return new FunctionNameEval(name);
            }
        }
        private ValueEval GetLocalNameXEval(NameXPtg nameXPtg)
        {
            String name = _workbook.ResolveNameXText(nameXPtg);

            // Try to parse it as a name
            int sheetNameAt = name.IndexOf('!');
            IEvaluationName evalName = null;
            if (sheetNameAt > -1)
            {
                // Sheet based name
                String sheetName = name.Substring(0, sheetNameAt);
                String nameName = name.Substring(sheetNameAt + 1);
                evalName = _workbook.GetName(nameName, _workbook.GetSheetIndex(sheetName));
            }
            else
            {
                // Workbook based name
                evalName = _workbook.GetName(name, -1);
            }

            if (evalName != null)
            {
                // Process it as a name
                return new ExternalNameEval(evalName);
            }
            else
            {
                // Must be an external function
                return new FunctionNameEval(name);
            }
        }

        // Fetch the workbook this refers to, and the name as defined with that
        private ValueEval GetExternalNameXEval(ExternalName externName, String workbookName)
        {
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
                    // Need to Evaluate the reference in the context of the other book
                    OperationEvaluationContext refWorkbookContext = new OperationEvaluationContext(
                            refWorkbookEvaluator, refWorkbookEvaluator.Workbook, -1, -1, -1, _tracker);

                    Ptg ptg = evaluationName.NameDefinition[0];
                    if (ptg is Ref3DPtg)
                    {
                        Ref3DPtg ref3D = (Ref3DPtg)ptg;
                        return refWorkbookContext.GetRef3DEval(ref3D);
                    }
                    else if (ptg is Ref3DPxg)
                    {
                        Ref3DPxg ref3D = (Ref3DPxg)ptg;
                        return refWorkbookContext.GetRef3DEval(ref3D);
                    }
                    else if (ptg is Area3DPtg)
                    {
                        Area3DPtg area3D = (Area3DPtg)ptg;
                        return refWorkbookContext.GetArea3DEval(area3D);
                    }
                    else if (ptg is Area3DPxg)
                    {
                        Area3DPxg area3D = (Area3DPxg)ptg;
                        return refWorkbookContext.GetArea3DEval(area3D);
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