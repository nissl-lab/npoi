/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/
/*
 * Created on May 8, 2005
 *
 */
namespace NPOI.SS.Formula.Eval
{
    using NPOI.SS.Formula;

    /**
     * Evaluation of 2D (Row+Column) and 3D (Sheet+Row+Column) areas
     */
    public interface AreaEval : TwoDEval, ThreeDEval
    {

        /**
         * returns the 0-based index of the first row in
         * this area.
         */
        int FirstRow { get; }

        /**
         * returns the 0-based index of the last row in
         * this area.
         */
        int LastRow { get; }

        /**
         * returns the 0-based index of the first col in
         * this area.
         */
        int FirstColumn { get; }

        /**
         * returns the 0-based index of the last col in
         * this area.
         */
        int LastColumn { get; }


        /**
         * returns true if the cell at row and col specified 
         * as absolute indexes in the sheet is contained in 
         * this area.
         * @param row
         * @param col
         */
        bool Contains(int row, int col);

        /**
         * returns true if the specified col is in range
         * @param col
         */
        bool ContainsColumn(int col);

        /**
         * returns true if the specified row is in range
         * @param row
         */
        bool ContainsRow(int row);
        /**
 * @return the ValueEval from within this area at the specified row and col index. Never
 * <code>null</code> (possibly {@link BlankEval}).  The specified indexes should be absolute
 * indexes in the sheet and not relative indexes within the area.
 */
        ValueEval GetAbsoluteValue(int row, int col);
        /**
         * @return the ValueEval from within this area at the specified relativeRowIndex and 
         * relativeColumnIndex. Never <c>null</c> (possibly {@link BlankEval}). The
         * specified indexes should relative to the top left corner of this area.  
         */
        ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex);

        /**
         * Creates an {@link AreaEval} offset by a relative amount from from the upper left cell
         * of this area
         */
        AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx);
    }
}