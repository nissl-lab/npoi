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

namespace NPOI.SS.Util.CellWalk
{
    /**
     * @author Roman Kashitsyn
     */
    public interface ICellWalkContext
    {

        /**
         * Returns ordinal number of cell in range.  Numeration starts
         * from top left cell and ends at bottom right cell. Here is a
         * brief example (number in cell is it's ordinal number):
         *
         * <table border="1">
         *   <tbody>
         *     <tr><td>1</td><td>2</td></tr>
         *     <tr><td>3</td><td>4</td></tr>
         *   </tbody>
         * </table>
         *
         * @return ordinal number of current cell
         */
        long OrdinalNumber { get; }

        /**
         * Returns number of current row.
         * @return number of current row
         */
        int RowNumber { get; }

        /**
         * Returns number of current column.
         * @return number of current column
         */
        int ColumnNumber { get; }
    }
}