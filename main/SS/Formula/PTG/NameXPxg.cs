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
    using System.Text;
    using NPOI.Util;

    /**
     * A Name, be that a Named Range or a Function / User Defined
     *  Function, Addressed in the HSSF External Sheet style.
     *  
     * <p>This is XSSF only, as it stores the sheet / book references
     *  in String form. The HSSF equivalent using indexes is {@link NameXPtg}</p>
     */
    [Serializable]
    public class NameXPxg : OperandPtg, Pxg
    {
        private int externalWorkbookNumber = -1;
        private String sheetName;
        private String nameName;

        public NameXPxg(int externalWorkbookNumber, String sheetName, String nameName)
        {
            this.externalWorkbookNumber = externalWorkbookNumber;
            this.sheetName = sheetName;
            this.nameName = nameName;
        }
        public NameXPxg(String sheetName, String nameName)
            : this(-1, sheetName, nameName)
        {
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
            if (SheetName != null)
            {
                sb.Append("sheet=").Append(SheetName);
            }
            sb.Append(" ! ");
            sb.Append("name=");
            sb.Append(nameName);
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
        public String NameName
        {
            get
            {
                return nameName;
            }
        }

        public override String ToFormulaString()
        {
            StringBuilder sb = new StringBuilder();
            bool needsExclamation = false;
            if (externalWorkbookNumber >= 0)
            {
                sb.Append('[');
                sb.Append(externalWorkbookNumber);
                sb.Append(']');
                needsExclamation = true;
            }
            if (sheetName != null)
            {
                SheetNameFormatter.AppendFormat(sb, sheetName);
                needsExclamation = true;
            }
            if (needsExclamation)
            {
                sb.Append('!');
            }
            sb.Append(nameName);
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