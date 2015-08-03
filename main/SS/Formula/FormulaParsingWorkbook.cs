/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula
{

    using System;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Util;

    /**
     * Abstracts a workbook for the purpose of formula parsing.<br/>
     * 
     * For POI internal use only
     * 
     * @author Josh Micich
     */
    public interface IFormulaParsingWorkbook
    {
        /// <summary>
        /// named range name matching is case insensitive
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <returns></returns>        
        IEvaluationName GetName(String name, int sheetIndex);

        /// <summary>
        /// Gets the name XPTG.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        Ptg GetNameXPtg(String name, SheetIdentifier sheet);

        /// <summary>
        /// Produce the appropriate Ptg for a 3d cell reference
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        Ptg Get3DReferencePtg(CellReference cell, SheetIdentifier sheet);

        /// <summary>
        /// Produce the appropriate Ptg for a 3d area reference
        /// </summary>
        /// <param name="area"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        Ptg Get3DReferencePtg(AreaReference area, SheetIdentifier sheet);

        /// <summary>
        /// Gets the externSheet index for a sheet from this workbook
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns></returns>
        int GetExternalSheetIndex(String sheetName);
        /// <summary>
        /// Gets the externSheet index for a sheet from an external workbook
        /// </summary>
        /// <param name="workbookName">Name of the workbook, e.g. "BudGet.xls"</param>
        /// <param name="sheetName">a name of a sheet in that workbook</param>
        /// <returns></returns>
        int GetExternalSheetIndex(String workbookName, String sheetName);

        /// <summary>
        /// Returns an enum holding spReadhseet properties specific to an Excel version (
        /// max column and row numbers, max arguments to a function, etc.)
        /// </summary>
        /// <returns></returns>
        SpreadsheetVersion GetSpreadsheetVersion();
    }
}