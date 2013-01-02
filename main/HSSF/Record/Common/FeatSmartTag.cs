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

namespace NPOI.HSSF.Record.Common
{
    using System;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.Util;

    /**
     * Title: FeatSmartTag (Smart Tag Shared Feature) common record part
     * 
     * This record part specifies Smart Tag data for a sheet, stored as part
     *  of a Shared Feature. It can be found in records such as  {@link FeatRecord}.
     * It is made up of a hash, and a Set of Factoid Data that Makes up
     *  the smart tags.
     * For more details, see page 669 of the Excel binary file
     *  format documentation.
     */
    public class FeatSmartTag : SharedFeature
    {
        // TODO - process
        private byte[] data;

        public FeatSmartTag()
        {
            data = new byte[0];
        }

        public FeatSmartTag(RecordInputStream in1)
        {
            data = in1.ReadRemainder();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append(" [FEATURE SMART TAGS]\n");
            buffer.Append(" [/FEATURE SMART TAGS]\n");
            return buffer.ToString();
        }

        public int DataSize
        {
            get
            {
                return data.Length;
            }
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            out1.Write(data);
        }
    }

}