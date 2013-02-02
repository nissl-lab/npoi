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

namespace NPOI.HSSF.Model
{
    using System;
    using NPOI.DDF;
    using NPOI.HSSF.UserModel;

    public class ConvertAnchor
    {
        /// <summary>
        /// Creates the anchor.
        /// </summary>
        /// <param name="userAnchor">The user anchor.</param>
        /// <returns></returns>
        public static EscherRecord CreateAnchor(HSSFAnchor userAnchor)
        {
            if (userAnchor is HSSFClientAnchor)
            {
                HSSFClientAnchor a = (HSSFClientAnchor)userAnchor;

                EscherClientAnchorRecord anchor = new EscherClientAnchorRecord();
                anchor.RecordId=EscherClientAnchorRecord.RECORD_ID;
                anchor.Options=(short)0x0000;
                anchor.Flag=(short)a.AnchorType;
                anchor.Col1=(short)Math.Min(a.Col1, a.Col2);
                anchor.Dx1=(short)a.Dx1;
                anchor.Row1=(short)Math.Min(a.Row1, a.Row2);
                anchor.Dy1=(short)a.Dy1;

                anchor.Col2=(short)Math.Max(a.Col1, a.Col2);
                anchor.Dx2=(short)a.Dx2;
                anchor.Row2=(short)Math.Max(a.Row1, a.Row2);
                anchor.Dy2=(short)a.Dy2;
                return anchor;
            }
            else
            {
                HSSFChildAnchor a = (HSSFChildAnchor)userAnchor;
                EscherChildAnchorRecord anchor = new EscherChildAnchorRecord();
                anchor.RecordId=EscherChildAnchorRecord.RECORD_ID;
                anchor.Options=(short)0x0000;
                anchor.Dx1=(short)Math.Min(a.Dx1, a.Dx2);
                anchor.Dy1=(short)Math.Min(a.Dy1, a.Dy2);
                anchor.Dx2=(short)Math.Max(a.Dx2, a.Dx1);
                anchor.Dy2=(short)Math.Max(a.Dy2, a.Dy1);
                return anchor;
            }
        }

    }
}