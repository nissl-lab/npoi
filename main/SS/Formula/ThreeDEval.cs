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

namespace NPOI.SS.Formula
{
    using System;
    using NPOI.SS.Formula.Eval;

    /// <summary>
    /// Optional Extension to the likes of <see cref="AreaEval"/> and
    ///  <see cref="AreaEvalBase" />,
    ///  which allows for looking up 3D (sheet+row+column) Evaluations
    /// </summary>
    public interface ThreeDEval : TwoDEval, ISheetRange
    {
        /// <summary>
        /// </summary>
        /// <param name="sheetIndex">sheet index (zero based)</param>
        /// <param name="rowIndex">relative row index (zero based)</param>
        /// <param name="columnIndex">relative column index (zero based)</param>
        /// <returns>element at the specified row and column position</returns>
        ValueEval GetValue(int sheetIndex, int rowIndex, int columnIndex);
    }

}