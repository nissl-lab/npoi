
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
     * Title:        Backup Record 
     * Description:  bool specifying whether
     *               the GUI should store a backup of the file.
     * REFERENCE:  PG 287 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class BackupRecord : StandardRecord
    {
        public const short sid = 0x40;
        private short field_1_backup;   // = 0;

        public BackupRecord()
        {
        }

        /**
         * Constructs a BackupRecord and Sets its fields appropriately
         * @param in the RecordInputstream to Read the record from
         */

        public BackupRecord(RecordInputStream in1)
        {
             field_1_backup = in1.ReadShort();
        }

        /**
         * Get the backup flag
         *
         * @return short 0/1 (off/on)
         */

        public short Backup
        {
            get { return field_1_backup; }
            set { field_1_backup = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[BACKUP]\n");
            buffer.Append("    .backup          = ")
                .Append(StringUtil.ToHexString(Backup)).Append("\n");
            buffer.Append("[/BACKUP]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(Backup);
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