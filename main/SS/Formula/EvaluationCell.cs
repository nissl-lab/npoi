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
    using NPOI.SS.UserModel;
    /**
     * Abstracts a cell for the purpose of formula evaluation.  This interface represents both formula
     * and non-formula cells.<br/>
     * 
     * Implementors of this class must implement {@link #HashCode()} and {@link #Equals(Object)}
     * To provide an <em>identity</em> relationship based on the underlying HSSF or XSSF cell <p/>
     * 
     * For POI internal use only
     * 
     * @author Josh Micich
     */
    public interface IEvaluationCell
    {
        // consider method Object GetUnderlyingCell() To reduce memory consumption in formula cell cache
        IEvaluationSheet Sheet { get; }
        int RowIndex { get; }
        int ColumnIndex { get; }
        CellType CellType { get; }

        double NumericCellValue { get; }
        String StringCellValue { get; }
        bool BooleanCellValue { get; }
        int ErrorCellValue { get; }
        Object IdentityKey { get; }
        CellType CachedFormulaResultType { get; }
    }
}