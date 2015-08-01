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
     * <p>Title:        XSSF Area 3D Reference (Sheet + Area)</p>
     * <p>Description:  Defined an area in an external or different sheet. </p>
     * <p>REFERENCE:  </p>
     * 
     * <p>This is XSSF only, as it stores the sheet / book references
     *  in String form. The HSSF equivalent using indexes is {@link Area3DPtg}</p>
     */
    public class Area3DPxg : AreaPtgBase, Pxg3D
    {
        private int externalWorkbookNumber = -1;
        private String firstSheetName;
        private String lastSheetName;

        public Area3DPxg(int externalWorkbookNumber, SheetIdentifier sheetName, String arearef)
            : this(externalWorkbookNumber, sheetName, new AreaReference(arearef))
        {
            ;
        }
        public Area3DPxg(int externalWorkbookNumber, SheetIdentifier sheetName, AreaReference arearef)
            : base(arearef)
        {
            this.externalWorkbookNumber = externalWorkbookNumber;
            this.firstSheetName = sheetName.SheetId.Name;
            if (sheetName is SheetRangeIdentifier)
            {
                this.lastSheetName = ((SheetRangeIdentifier)sheetName).LastSheetIdentifier.Name;
            }
            else
            {
                this.lastSheetName = null;
            }
        }

        public Area3DPxg(SheetIdentifier sheetName, String arearef)
            : this(sheetName, new AreaReference(arearef))
        {
        }
        public Area3DPxg(SheetIdentifier sheetName, AreaReference arearef)
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
            sb.Append("sheet=").Append(SheetName);
            if (lastSheetName != null)
            {
                sb.Append(" : ");
                sb.Append("sheet=").Append(lastSheetName);
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
                return firstSheetName;
            }
            set
            {
                firstSheetName = value;
            }
        }

        public string LastSheetName
        {
            get { return lastSheetName; }
            set { lastSheetName = value; }
        }

        public String Format2DRefAsString()
        {
            return FormatReferenceAsString();
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
            SheetNameFormatter.AppendFormat(sb, firstSheetName);
            if (lastSheetName != null)
            {
                sb.Append(':');
                SheetNameFormatter.AppendFormat(sb, lastSheetName);
            }
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

