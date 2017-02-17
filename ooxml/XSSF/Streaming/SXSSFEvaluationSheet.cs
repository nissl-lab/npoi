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
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFEvaluationSheet : XSSFEvaluationSheet
    {
        private SXSSFSheet _xs;

    public SXSSFEvaluationSheet(SXSSFSheet sheet)
        {
            _xs = sheet;
        }

        public SXSSFSheet getSXSSFSheet()
        {
            return _xs;
        }

    public SXSSFEvaluationCell getCell(int rowIndex, int columnIndex)
        {
            SXSSFRow row = _xs._rows[rowIndex];
            if (row == null)
            {
                if (rowIndex <= _xs.lastFlushedRowNumber)
                {
                    throw new RowFlushedException(rowIndex);
                }
                return null;
            }
            SXSSFCell cell = (SXSSFCell)row.Cells[columnIndex];
            if (cell == null)
            {
                return null;
            }
            return new SXSSFEvaluationCell(cell, this);
        }

        /* (non-JavaDoc), inherit JavaDoc from EvaluationSheet
         * @since POI 3.15 beta 3
         */

    public void clearAllCachedResultValues()
        {
            // nothing to do
        }
    }
}
