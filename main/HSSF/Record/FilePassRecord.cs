
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
     * Title:        File Pass Record
     * Description:  Indicates that the record after this record are encrypted. HSSF does not support encrypted excel workbooks
     * and the presence of this record will cause Processing to be aborted.
     * REFERENCE:  PG 420 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 3.0-pre
     */

    public class FilePassRecord
       : StandardRecord
    {
        public const short sid = 0x2F;
        private int _encryptionType;
        private int _encryptionInfo;
        private int _minorVersionNo;
        private byte[] _docId;
        private byte[] _saltData;
        private byte[] _saltHash;


        private const int ENCRYPTION_XOR = 0;
        private const int ENCRYPTION_OTHER = 1;

        private const int ENCRYPTION_OTHER_RC4 = 1;
        private const int ENCRYPTION_OTHER_CAPI_2 = 2;
        private const int ENCRYPTION_OTHER_CAPI_3 = 3;

        public FilePassRecord()
        {
        }

        public FilePassRecord(RecordInputStream in1)
        {
            _encryptionType = in1.ReadUShort();

            switch (_encryptionType)
            {
                case ENCRYPTION_XOR:
                    throw new RecordFormatException("HSSF does not currently support XOR obfuscation");
                case ENCRYPTION_OTHER:
                    // handled below
                    break;
                default:
                    throw new RecordFormatException("Unknown encryption type " + _encryptionType);
            }
            _encryptionInfo = in1.ReadUShort();
            switch (_encryptionInfo)
            {
                case ENCRYPTION_OTHER_RC4:
                    // handled below
                    break;
                case ENCRYPTION_OTHER_CAPI_2:
                case ENCRYPTION_OTHER_CAPI_3:
                    throw new RecordFormatException(
                            "HSSF does not currently support CryptoAPI encryption");
                default:
                    throw new RecordFormatException("Unknown encryption info " + _encryptionInfo);
            }
            _minorVersionNo = in1.ReadUShort();
            if (_minorVersionNo != 1)
            {
                throw new RecordFormatException("Unexpected VersionInfo number for RC4Header " + _minorVersionNo);
            }
            _docId = Read(in1, 16);
            _saltData = Read(in1, 16);
            _saltHash = Read(in1, 16);
        }
        private static byte[] Read(RecordInputStream in1, int size)
        {
            byte[] result = new byte[size];
            in1.ReadFully(result);
            return result;
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FILEPASS]\n");
            buffer.Append("    .type = ").Append(HexDump.ShortToHex(_encryptionType)).Append("\n");
            buffer.Append("    .info = ").Append(HexDump.ShortToHex(_encryptionInfo)).Append("\n");
            buffer.Append("    .ver  = ").Append(HexDump.ShortToHex(_minorVersionNo)).Append("\n");
            buffer.Append("    .docId= ").Append(HexDump.ToHex(_docId)).Append("\n");
            buffer.Append("    .salt = ").Append(HexDump.ToHex(_saltData)).Append("\n");
            buffer.Append("    .hash = ").Append(HexDump.ToHex(_saltHash)).Append("\n");
            buffer.Append("[/FILEPASS]\n");
            return buffer.ToString();
        }


        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(_encryptionType);
            out1.WriteShort(_encryptionInfo);
            out1.WriteShort(_minorVersionNo);
            out1.Write(_docId);
            out1.Write(_saltData);
            out1.Write(_saltHash);
        }

        protected override int DataSize
        {
            get
            {
                return 54;
            }
        }
        public byte[] DocId
        {
            get
            {
                return _docId;
            }
            set
            {

            }
        }

        public byte[] SaltData
        {
            get
            {
                return _saltData;
            }
        }

        public byte[] SaltHash
        {
            get
            {
                return _saltHash;
            }
        }


        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            // currently immutable
            return this;
        }
    }
}