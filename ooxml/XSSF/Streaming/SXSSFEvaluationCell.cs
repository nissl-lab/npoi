/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
using System;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFEvaluationCell : XSSFEvaluationCell
    {
    private  SXSSFEvaluationSheet _evalSheet;
    private  SXSSFCell _cell;

    public SXSSFEvaluationCell(SXSSFCell cell, SXSSFEvaluationSheet evaluationSheet)
    {
        _cell = cell;
        _evalSheet = evaluationSheet;
    }

    public SXSSFEvaluationCell(SXSSFCell cell): this(cell, null/*new SXSSFEvaluationSheet(cell.Sheet)*/)
    {
        throw new NotImplementedException();
    }


    public Object getIdentityKey()
    {
        // save memory by just using the cell itself as the identity key
        // Note - this assumes SXSSFCell has not overridden hashCode and equals
        return _cell;
    }

    public SXSSFCell getSXSSFCell()
    {
        return _cell;
    }
  
    public bool getBooleanCellValue()
    {
        return _cell.BooleanCellValue;
    }
    /**
     * Will return {@link CellType} in a future version of POI.
     * For forwards compatibility, do not hard-code cell type literals in your code.
     *
     * @return cell type
     */

    public int getCellType()
    {
        return (int)_cell.CellType;
    }
    /**
     * @since POI 3.15 beta 3
     * @deprecated POI 3.15 beta 3.
     * Will be deleted when we make the CellType enum transition. See bug 59791.
     */
    //(since= "POI 3.15 beta 3")
   
    public CellType getCellTypeEnum()
    {
        return _cell.CellType;
    }

    public int getColumnIndex()
    {
        return _cell.ColumnIndex;
    }
    
    public int getErrorCellValue()
    {
        return _cell.ErrorCellValue;
    }
    
    public double getNumericCellValue()
    {
        return _cell.NumericCellValue;
    }
    
    public int getRowIndex()
    {
        return _cell.RowIndex;
    }
    
    public SXSSFEvaluationSheet getSheet()
    {
        return _evalSheet;
    }
    
    public String getStringCellValue()
    {
        return _cell.RichStringCellValue.String;
    }
    /**
     * Will return {@link CellType} in a future version of POI.
     * For forwards compatibility, do not hard-code cell type literals in your code.
     *
     * @return cell type of cached formula result
     */
    
    public int getCachedFormulaResultType()
    {
        return (int)_cell.CachedFormulaResultType;
    }
    /**
     * @since POI 3.15 beta 3
     * @deprecated POI 3.15 beta 3.
     * Will be deleted when we make the CellType enum transition. See bug 59791.
     */
    //@Internal(since= "POI 3.15 beta 3")
    
    public CellType getCachedFormulaResultTypeEnum()
    {
        return _cell.GetCachedFormulaResultTypeEnum();
    }
}
}
