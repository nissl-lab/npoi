using NPOI.SS.Formula.Eval;
using System;

namespace NPOI.SS.Formula.Functions
{
    /// <summary>
    /// Implementation for Excel SHEET() function.
    /// </summary>
    public class Sheet : FreeRefFunction
    {

        public static Sheet instance = new Sheet();

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            try
            {
                if(args.Length == 0)
                {
                    // No argument provided → return the current sheet index +1 (Excel uses 1-based index)
                    return new NumberEval((double) ec.SheetIndex + 1);
                }
                else
                {
                    ValueEval arg = args[0];

                    if(arg is RefEval)
                    {
                        // Argument is a single cell reference → return the sheet index of that reference +1
                        var ref1 = (RefEval) arg;
                        int sheetIndex = ref1.FirstSheetIndex;
                        return new NumberEval((double) sheetIndex + 1);
                    }
                    else if(arg is AreaEval)
                    {
                        // Argument is a cell range → return the sheet index of that area +1
                        AreaEval area = (AreaEval) arg;
                        int sheetIndex = area.FirstSheetIndex;
                        return new NumberEval((double) sheetIndex + 1);
                    }
                    else if(arg is StringEval)
                    {
                        // Argument is a string (sheet name, e.g., "Sheet3") → look up the sheet index by name
                        String sheetName = ((StringEval) arg).StringValue;
                        var wb = ec.GetWorkbook();
                        int sheetIndex = wb.GetSheetIndex(sheetName);
                        if(sheetIndex >= 0)
                        {
                            return new NumberEval((double) sheetIndex + 1);
                        }
                        else
                        {
                            // Sheet name not found → return #N/A error
                            return ErrorEval.NA;
                        }
                    }
                    else
                    {
                        // Unsupported argument type → return #N/A error
                        return ErrorEval.NA;
                    }
                }
            }
            catch(Exception e)
            {
                // Any unexpected exception (e.g., null pointers) → return #VALUE! error
                return ErrorEval.VALUE_INVALID;
            }
        }
    }
}
