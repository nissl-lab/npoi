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

    /**
     * Common interface of {@link AreaEval} and {@link org.apache.poi.ss.formula.eval.AreaEvalBase},
     * for 2D (row+column) evaluations
     */
    public interface TwoDEval : ValueEval
    {

        /**
         * @param rowIndex relative row index (zero based)
         * @param columnIndex relative column index (zero based)
         * @return element at the specified row and column position
         */
        ValueEval GetValue(int rowIndex, int columnIndex);

        int Width { get; }
        int Height { get; }

        /**
         * @return <c>true</c> if the area has just a single row, this also includes
         * the trivial case when the area has just a single cell.
         */
        bool IsRow { get; }

        /**
         * @return <c>true</c> if the area has just a single column, this also includes
         * the trivial case when the area has just a single cell.
         */
        bool IsColumn { get; }

        /**
         * @param rowIndex relative row index (zero based)
         * @return a single row {@link TwoDEval}
         */
        TwoDEval GetRow(int rowIndex);
        /**
         * @param columnIndex relative column index (zero based)
         * @return a single column {@link TwoDEval}
         */
        TwoDEval GetColumn(int columnIndex);


        /**
         * @return true if the  cell at row and col is a subtotal
         */
        bool IsSubTotal(int rowIndex, int columnIndex);

    }

}





