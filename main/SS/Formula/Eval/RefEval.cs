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

namespace NPOI.SS.Formula.Eval
{
    /**
     * @author Amol S Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * 
     * RefEval is the base interface for Ref2D and Ref3DEval. Basically a RefEval
     * impl should contain reference to the original ReferencePtg or Ref3DPtg as
     * well as the "value" resulting from the evaluation of the cell
     * reference. Thus if the HSSFCell has type CELL_TYPE_NUMERIC, the contained
     * value object should be of type NumberEval; if cell type is CELL_TYPE_STRING,
     * contained value object should be of type StringEval
     */
    public interface RefEval : ValueEval, ISheetRange
    {

        /**
         * The (possibly Evaluated) ValueEval contained
         * in this RefEval. eg. if cell A1 Contains "test"
         * then in a formula referring to cell A1 
         * the RefEval representing
         * A1 will return as the InnerValueEval the
         * object of concrete type StringEval
         */
        ValueEval GetInnerValueEval(int sheetIndex);
        /**
         * returns the zero based column index.
         */
        int Column { get; }

        /**
         * returns the zero based row index.
         */
        int Row { get; }

        /**
         * returns the first sheet index this applies to
         */
        int FirstSheetIndex { get; }

        /**
         * returns the last sheet index this applies to, which
         *  will be the same as the first for a 2D and many 3D references
         */
        int LastSheetIndex { get; }

        /**
         * returns the number of sheets this applies to
         */
        int NumberOfSheets { get; }

        /**
         * Creates an {@link AreaEval} offset by a relative amount from this RefEval
         */
        AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx);

    }
}