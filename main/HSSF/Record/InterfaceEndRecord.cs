
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


namespace NPOI.HSSF.Record
{
    using System;
    using NPOI.Util;


    /**
     * Title: Interface End Record
     * Description: Shows where the Interface Records end (MMS)
     *  (has no fields)
     * REFERENCE:  PG 324 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class InterfaceEndRecord
       : StandardRecord
    {
        public const short sid = 0xe2;
        public static InterfaceEndRecord Instance = new InterfaceEndRecord();
        private InterfaceEndRecord()
        {
            // enforce singleton
        }

        /**
         * Constructs an InterfaceEnd record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */
        //private byte[] _unknownData;


        //public InterfaceEndRecord(RecordInputStream in1)
        //{
        //    if(in1.Available() > 0){
        //        _unknownData = in1.ReadRemainder();
        //    }

        //}
        public static Record Create(RecordInputStream in1)
        {
            switch (in1.Remaining)
            {
                case 0:
                    return Instance;
                case 2:
                    return new InterfaceHdrRecord(in1);
            }
            throw new RecordFormatException("Invalid record data size: " + in1.Remaining);
        }

        public override String ToString()
        {
            return "[INTERFACEEND/]\n";
            //StringBuilder buffer = new StringBuilder();

            //buffer.Append("[INTERFACEEND]\n");
            //buffer.Append("  unknownData=").Append(HexDump.ToHex(_unknownData)).Append("\n");
            //buffer.Append("[/INTERFACEEND]\n");
            //return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            //if(_unknownData != null) out1.Write(_unknownData);
            // no instance data
        }
        /// <summary>
        /// for test TestInterfaceEndRecord.TestCreate()
        /// </summary>
        /// <returns></returns>
        public int GetDataSize()
        {
            return DataSize;
        }
        protected override int DataSize
        {
            get 
            {
                int size = 0;
                //if (_unknownData != null) size += _unknownData.Length;

                return size;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}
