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

using NPOI.SS.Formula.Eval;
namespace NPOI.SS.Formula
{

    /// <summary>
    /// Common interface of <see cref="AreaEval"/> and <see cref="org.apache.poi.ss.formula.eval.AreaEvalBase" />,
    /// for 2D (row+column) evaluations
    /// </summary>
    public interface TwoDEval : ValueEval
    {

        /// <summary>
        /// </summary>
        /// <param name="rowIndex">relative row index (zero based)</param>
        /// <param name="columnIndex">relative column index (zero based)</param>
        /// <returns>element at the specified row and column position</returns>
        ValueEval GetValue(int rowIndex, int columnIndex);

        int Width { get; }
        int Height { get; }

        /// <summary>
        /// </summary>
        /// <returns><c>true</c> if the area has just a single row, this also includes
        /// the trivial case when the area has just a single cell.
        /// </returns>
        bool IsRow { get; }

        /// <summary>
        /// </summary>
        /// <returns><c>true</c> if the area has just a single column, this also includes
        /// the trivial case when the area has just a single cell.
        /// </returns>
        bool IsColumn { get; }

        /// <summary>
        /// </summary>
        /// <param name="rowIndex">relative row index (zero based)</param>
        /// <returns>a single row <see cref="TwoDEval"/></returns>
        TwoDEval GetRow(int rowIndex);
        /// <summary>
        /// </summary>
        /// <param name="columnIndex">relative column index (zero based)</param>
        /// <returns>a single column <see cref="TwoDEval"/></returns>
        TwoDEval GetColumn(int columnIndex);


        /// <summary>
        /// </summary>
        /// <returns>true if the  cell at row and col is a subtotal</returns>
        bool IsSubTotal(int rowIndex, int columnIndex);

    }

}
