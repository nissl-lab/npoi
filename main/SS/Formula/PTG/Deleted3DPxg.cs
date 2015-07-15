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

namespace NPOI.SS.Formula.PTG
{
    using System;

    using NPOI.SS.UserModel;
    using System.Text;
    using NPOI.Util;


    /**
     * An XSSF only representation of a reference to a deleted area
     */
    public class Deleted3DPxg : OperandPtg, Pxg
    {
        private int externalWorkbookNumber = -1;
        private String sheetName;

        public Deleted3DPxg(int externalWorkbookNumber, String sheetName)
        {
            this.externalWorkbookNumber = externalWorkbookNumber;
            this.sheetName = sheetName;
        }
        public Deleted3DPxg(String sheetName)
            : this(-1, sheetName)
        {
            ;
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(GetType().Name);
            sb.Append(" [");
            if (externalWorkbookNumber >= 0)
            {
                sb.Append(" [");
                sb.Append("workbook=").Append(ExternalWorkbookNumber);
                sb.Append("] ");
            }
            if (sheetName != null)
            {
                SheetNameFormatter.AppendFormat(sb, sheetName);
            }
            sb.Append(" ! ");
            sb.Append(ErrorConstants.GetText(ErrorConstants.ERROR_REF));
            sb.Append("]");
            return sb.ToString();
        }

        public int ExternalWorkbookNumber
        {
            get
            {
                return externalWorkbookNumber;
            }
        }
        public String SheetName
        {
            get
            {
                return sheetName;
            }
            set
            {
                sheetName = value;
            }
        }

        public override String ToFormulaString()
        {
            StringBuilder sb = new StringBuilder();
            if (externalWorkbookNumber >= 0)
            {
                sb.Append('[');
                sb.Append(externalWorkbookNumber);
                sb.Append(']');
            }
            if (sheetName != null)
            {
                sb.Append(sheetName);
            }
            sb.Append('!');
            sb.Append(ErrorConstants.GetText(ErrorConstants.ERROR_REF));
            return sb.ToString();
        }

        public override byte DefaultOperandClass
        {
            get
            {
                return Ptg.CLASS_VALUE;
            }
        }

        public override int Size
        {
            get
            {
                return 1;
            }
        }
        public override void Write(ILittleEndianOutput out1)
        {
            throw new InvalidOperationException("XSSF-only Ptg, should not be serialised");
        }
    }

}