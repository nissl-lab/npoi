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

namespace NPOI.HSSF.Record.CF
{
    using NPOI.Util;
    using System;

    /**
     * Icon / Multi-State specific Threshold / value (CFVO),
     *  for Changes in Conditional Formatting
     */
    public class IconMultiStateThreshold : Threshold, ICloneable
    {
        /**
         * Cell values that are equal to the threshold value do not pass the threshold
         */
        public static byte EQUALS_EXCLUDE = 0;
        /**
         * Cell values that are equal to the threshold value pass the threshold.
         */
        public static byte EQUALS_INCLUDE = 1;

        private byte equals;

        public IconMultiStateThreshold() : base()
        {

            equals = EQUALS_INCLUDE;
        }

        /** Creates new Ico Multi-State Threshold */
        public IconMultiStateThreshold(ILittleEndianInput in1)
                : base(in1)
        {

            equals = (byte)in1.ReadByte();
            // Reserved, 4 bytes, all 0
            in1.ReadInt();
        }

        public byte GetEquals()
        {
            return equals;
        }
        public void SetEquals(byte Equals)
        {
            this.equals = Equals;
        }

        public override int DataLength
        {

            get
            {
                return base.DataLength + 5;
            }


        }

        public Object Clone()
        {
            IconMultiStateThreshold rec = new IconMultiStateThreshold();
            base.CopyTo(rec);
            rec.equals = equals;
            return rec;
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            base.Serialize(out1);
            out1.WriteByte(equals);
            out1.WriteInt(0); // Reserved
        }
    }
}

