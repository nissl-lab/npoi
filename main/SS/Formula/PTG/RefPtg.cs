/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.SS.Formula.PTG
{
    using System;
    using NPOI.Util;
    using NPOI.SS.Util;

    /**
     * ReferencePtg - handles references (such as A1, A2, IA4)
     * @author  Andrew C. Oliver (acoliver@apache.org)
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    [Serializable]
    public class RefPtg : Ref2DPtgBase
    {
        public const byte sid = 0x24;

        /**
         * Takes in a String representation of a cell reference and Fills out the
         * numeric fields.
         */
        public RefPtg(String cellref)
            : base(new CellReference(cellref))
        {

        }

        public RefPtg(int row, int column, bool isRowRelative, bool isColumnRelative)
            :base(row, column, isRowRelative, isColumnRelative)
        {
            Row = row;
            Column = column;
            IsRowRelative = isRowRelative;
            IsColRelative = isColumnRelative;
        }

        public RefPtg(ILittleEndianInput in1)
            : base(in1)
        {

        }
        public RefPtg(CellReference cr):base(cr)
        {
            
        }
        protected override byte Sid
        {
            get { return sid; }
        }

    }
}