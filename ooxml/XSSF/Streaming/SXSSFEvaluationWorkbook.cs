using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.Formula;
using NPOI.SS.Formula.PTG;
using NPOI.XSSF.Streaminging;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public  class SXSSFEvaluationWorkbook : XSSFEvaluationWorkbook
    {
    private  SXSSFWorkbook _uBook;
    
    public static SXSSFEvaluationWorkbook create(SXSSFWorkbook book)
    {
        if (book == null)
        {
            return null;
        }
        return new SXSSFEvaluationWorkbook(book);
    }

    private SXSSFEvaluationWorkbook(SXSSFWorkbook book) : base(book.XssfWorkbook)
    {
        _uBook = book;
    }

    
    public int getSheetIndex(SXSSFEvaluationSheet evalSheet)
    {
        SXSSFSheet sheet = ((SXSSFEvaluationSheet)evalSheet).getSXSSFSheet();
        return _uBook.GetSheetIndex(sheet);
    }

    
    public SXSSFEvaluationSheet getSheet(int sheetIndex)
    {
            throw new NotImplementedException();
        //return new SXSSFEvaluationSheet(_uBook.GetSheetAt(sheetIndex));
    }

    
    public Ptg[] getFormulaTokens(SXSSFEvaluationCell evalCell)
    {
        SXSSFCell cell = ((SXSSFEvaluationCell)evalCell).getSXSSFCell();
        return FormulaParser.Parse(cell.CellFormula, this, FormulaType.Cell, _uBook.GetSheetIndex(cell.Sheet));
    }
}

}
