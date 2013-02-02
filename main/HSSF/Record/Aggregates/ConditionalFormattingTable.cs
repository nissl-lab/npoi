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

namespace NPOI.HSSF.Record.Aggregates
{

    using System;
    using System.Collections;

    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;

    /**
     * Holds all the conditional formatting for a workbook sheet.<p/>
     * 
     * See OOO exelfileformat.pdf sec 4.12 'Conditional Formatting Table'
     * 
     * @author Josh Micich
     */
    public class ConditionalFormattingTable : RecordAggregate
    {

        private IList _cfHeaders;

        /**
         * Creates an empty ConditionalFormattingTable
         */
        public ConditionalFormattingTable()
        {
            _cfHeaders = new ArrayList();
        }

        public ConditionalFormattingTable(RecordStream rs)
        {

            IList temp = new ArrayList();
            while (rs.PeekNextClass() == typeof(CFHeaderRecord))
            {
                temp.Add(CFRecordsAggregate.CreateCFAggregate(rs));
            }
            _cfHeaders = temp;
        }

        public override void VisitContainedRecords(RecordVisitor rv)
        {
            for (int i = 0; i < _cfHeaders.Count; i++)
            {
                CFRecordsAggregate subAgg = (CFRecordsAggregate)_cfHeaders[i];
                subAgg.VisitContainedRecords(rv);
            }
        }

        /**
         * @return index of the newly added CF header aggregate
         */
        public int Add(CFRecordsAggregate cfAggregate)
        {
            _cfHeaders.Add(cfAggregate);
            return _cfHeaders.Count - 1;
        }

        public int Count
        {
            get
            {
                return _cfHeaders.Count;
            }
        }

        public CFRecordsAggregate Get(int index)
        {
            CheckIndex(index);
            return (CFRecordsAggregate)_cfHeaders[index];
        }

        public void Remove(int index)
        {
            CheckIndex(index);
            _cfHeaders.RemoveAt(index);
        }

        private void CheckIndex(int index)
        {
            if (index < 0 || index >= _cfHeaders.Count)
            {
                throw new ArgumentException("Specified CF index " + index
                        + " is outside the allowable range (0.." + (_cfHeaders.Count - 1) + ")");
            }
        }

        public void UpdateFormulasAfterCellShift(FormulaShifter shifter, int externSheetIndex)
        {
            for (int i = 0; i < _cfHeaders.Count; i++)
            {
                CFRecordsAggregate subAgg = (CFRecordsAggregate)_cfHeaders[i];
                bool shouldKeep = subAgg.UpdateFormulasAfterCellShift(shifter, externSheetIndex);
                if (!shouldKeep)
                {
                    _cfHeaders.RemoveAt(i);
                    i--;
                }
            }
        }
    }
}