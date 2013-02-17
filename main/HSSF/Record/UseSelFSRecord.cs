
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
    using System.Text;
    using NPOI.Util;

    /**
     * Title:        Use Natural Language Formulas Flag
     * Description:  Tells the GUI if this was written by something that can use
     *               "natural language" formulas. HSSF can't.
     * REFERENCE:  PG 420 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class UseSelFSRecord
       : StandardRecord
    {
        public const short sid = 0x160;

        private static BitField useNaturalLanguageFormulasFlag = BitFieldFactory.GetInstance(0x0001);
        private int _options;

        public UseSelFSRecord(int options)
        {
            _options = options;
        }

        /**
         * Constructs a UseSelFS record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public UseSelFSRecord(RecordInputStream in1) :
            this(in1.ReadUShort())
        {
        }
        public UseSelFSRecord(bool b)
            : this(0)
        {
            _options = useNaturalLanguageFormulasFlag.SetBoolean(_options, b);
        }
        // /**
        // * turn the flag on or off
        // *
        // * @param flag  whether to use natural language formulas or not
        // * @see #TRUE
        // * @see #FALSE
        // */

        //public void SetFlag(short flag)
        //{
        //    field_1_flag = flag;
        //}

        // /**
        // * returns whether we use natural language formulas or not
        // *
        // * @return whether to use natural language formulas or not
        // * @see #TRUE
        // * @see #FALSE
        // */

        //public short GetFlag()
        //{
        //    return field_1_flag;
        //}

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[USESELFS]\n");
            buffer.Append("    .flag            = ")
                .Append(HexDump.ShortToHex(_options)).Append("\n");
            buffer.Append("[/USESELFS]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(_options);
        }

        protected override int DataSize
        {
            get
            {
                return 2;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}