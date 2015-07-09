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
    using NPOI.SS.Util;
    using NPOI.Util;

/**
 * <p>Title:        XSSF Area 3D Reference (Sheet + Area)<P>
 * <p>Description:  Defined an area in an external or different sheet. <P>
 * <p>REFERENCE:  </p>
 * 
 * <p>This is XSSF only, as it stores the sheet / book references
 *  in String form. The HSSF equivalent using indexes is {@link Area3DPtg}</p>
 */
    public class Area3DPxg : AreaPtgBase, Pxg
    {
        private int externalWorkbookNumber = -1;
        private String sheetName;

        public Area3DPxg(int externalWorkbookNumber, String sheetName, String arearef)
            : this(externalWorkbookNumber, sheetName, new AreaReference(arearef))
        {
            ;
        }
        public Area3DPxg(int externalWorkbookNumber, String sheetName, AreaReference arearef)
            : base(arearef)
        {

            this.externalWorkbookNumber = externalWorkbookNumber;
            this.sheetName = sheetName;
        }

        public Area3DPxg(String sheetName, String arearef)
            : this(sheetName, new AreaReference(arearef))
        {
        }
        public Area3DPxg(String sheetName, AreaReference arearef)
            : this(-1, sheetName, arearef)
        {

        }


        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(this.GetType().Name);
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
            sb.Append(FormatReferenceAsString());
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

        public String Format2DRefAsString()
        {
            return FormatReferenceAsString();
        }

        public String toFormulaString()
        {
            StringBuilder sb = new StringBuilder();
            if (externalWorkbookNumber >= 0)
            {
                sb.Append('[');
                sb.Append(externalWorkbookNumber);
                sb.Append(']');
            }
            sb.Append(sheetName);
            sb.Append('!');
            sb.Append(FormatReferenceAsString());
            return sb.ToString();
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

