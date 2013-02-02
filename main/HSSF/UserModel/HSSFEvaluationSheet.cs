/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.HSSF.UserModel
{
    using NPOI.SS.Formula;
    using NPOI.SS.UserModel;

    /**
     * HSSF wrapper for a sheet under evaluation
     * 
     * @author Josh Micich
     */
    public class HSSFEvaluationSheet : IEvaluationSheet
    {

        private HSSFSheet _hs;

        public HSSFEvaluationSheet(HSSFSheet hs)
        {
            _hs = hs;
        }

        public HSSFSheet HSSFSheet
        {
            get
            {
                return _hs;
            }
        }
        public IEvaluationCell GetCell(int rowIndex, int columnIndex)
        {
            HSSFRow row = (HSSFRow)_hs.GetRow(rowIndex);
            if (row == null)
            {
                return null;
            }
            ICell cell = (HSSFCell)row.GetCell(columnIndex);
            if (cell == null)
            {
                return null;
            }
            return new HSSFEvaluationCell(cell, this);
        }
    }
}