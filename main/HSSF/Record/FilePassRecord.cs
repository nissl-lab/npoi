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

namespace NPOI.HSSF.Record
{
    using System;
    using NPOI.Util;
    using System.Text;
    using System.Diagnostics;

    /**
     * Title: File Pass Record (0x002F) <p>
     *
     * Description: Indicates that the record After this record are encrypted.
     */
    public class FilePassRecord : StandardRecord, ICloneable
    {
        public const short sid = 0x002F;
        private const int ENCRYPTION_XOR = 0;
        private const int ENCRYPTION_OTHER = 1;

        private readonly int _encryptionType;
        private readonly KeyData _keyData;

        private interface KeyData : ICloneable
        {
            void Read(RecordInputStream in1);
            void Serialize(ILittleEndianOutput out1);
            int DataSize { get; }
            void AppendToString(StringBuilder buffer);
            //object Clone(); // return KeyData;
        }

        public class Rc4KeyData : KeyData, ICloneable
        {
            private const int ENCRYPTION_OTHER_RC4 = 1;
            private const int ENCRYPTION_OTHER_CAPI_2 = 2;
            private const int ENCRYPTION_OTHER_CAPI_3 = 3;
            private const int ENCRYPTION_OTHER_CAPI_4 = 4;

            private byte[] _salt;
            private byte[] _encryptedVerifier;
            private byte[] _encryptedVerifierHash;
            private int _encryptionInfo;
            private int _minorVersionNo;

            public void Read(RecordInputStream in1)
            {
                _encryptionInfo = in1.ReadUShort();
                switch (_encryptionInfo)
                {
                    case ENCRYPTION_OTHER_RC4:
                        // handled below
                        break;
                    case ENCRYPTION_OTHER_CAPI_2:
                    case ENCRYPTION_OTHER_CAPI_3:
                    case ENCRYPTION_OTHER_CAPI_4:
                        throw new EncryptedDocumentException(
                                "HSSF does not currently support CryptoAPI encryption");
                    default:
                        throw new RecordFormatException("Unknown encryption info " + _encryptionInfo);
                }
                _minorVersionNo = in1.ReadUShort();
                if (_minorVersionNo != 1)
                {
                    throw new RecordFormatException("Unexpected VersionInfo number for RC4Header " + _minorVersionNo);
                }
                _salt = FilePassRecord.Read(in1, 16);
                _encryptedVerifier = FilePassRecord.Read(in1, 16);
                _encryptedVerifierHash = FilePassRecord.Read(in1, 16);
            }

            public void Serialize(ILittleEndianOutput out1)
            {
                out1.WriteShort(_encryptionInfo);
                out1.WriteShort(_minorVersionNo);
                out1.Write(_salt);
                out1.Write(_encryptedVerifier);
                out1.Write(_encryptedVerifierHash);
            }

            public int DataSize
            {
                get
                {
                    return 54;
                }
            }

            public byte[] Salt
            {
                get
                {
                    return (byte[])_salt.Clone();
                }
                set
                {
                    this._salt = (byte[])value.Clone();
                }
            }

            public byte[] EncryptedVerifier
            {
                get
                {
                    return (byte[])_encryptedVerifier.Clone();
                }
                set
                {
                    this._encryptedVerifier = (byte[])value.Clone();
                }
            }

            public byte[] EncryptedVerifierHash
            {
                get
                {
                    return (byte[])_encryptedVerifierHash.Clone();
                }
                set
                {
                    this._encryptedVerifierHash = (byte[])value.Clone();
                }
            }

            public void AppendToString(StringBuilder buffer)
            {
                buffer.Append("    .rc4.info = ").Append(HexDump.ShortToHex(_encryptionInfo)).Append("\n");
                buffer.Append("    .rc4.ver  = ").Append(HexDump.ShortToHex(_minorVersionNo)).Append("\n");
                buffer.Append("    .rc4.salt = ").Append(HexDump.ToHex(_salt)).Append("\n");
                buffer.Append("    .rc4.verifier = ").Append(HexDump.ToHex(_encryptedVerifier)).Append("\n");
                buffer.Append("    .rc4.verifierHash = ").Append(HexDump.ToHex(_encryptedVerifierHash)).Append("\n");
            }

            /// <summary>
            /// 
            /// </summary>
            /// <returns>Rc4KeyData</returns>
            public object Clone()
            {
                Rc4KeyData other = new Rc4KeyData();
                other._salt = (byte[])this._salt.Clone();
                other._encryptedVerifier = (byte[])this._encryptedVerifier.Clone();
                other._encryptedVerifierHash = (byte[])this._encryptedVerifierHash.Clone();
                other._encryptionInfo = this._encryptionInfo;
                other._minorVersionNo = this._minorVersionNo;
                return other;
            }
        }

        public class XorKeyData : KeyData, ICloneable
        {
            /**
             * key (2 bytes): An unsigned integer that specifies the obfuscation key. 
             * See [MS-OFFCRYPTO], 2.3.6.2 section, the first step of Initializing XOR
             * array where it describes the generation of 16-bit XorKey value.
             */
            private int _key;

            /**
             * verificationBytes (2 bytes): An unsigned integer that specifies
             * the password verification identifier.
             */
            private int _verifier;

            public void Read(RecordInputStream in1)
            {
                _key = in1.ReadUShort();
                _verifier = in1.ReadUShort();
            }

            public void Serialize(ILittleEndianOutput out1)
            {
                out1.WriteShort(_key);
                out1.WriteShort(_verifier);
            }

            public int DataSize
            {
                get
                {
                    // TODO: Check!
                    return 6;
                }
            }

            public int Key
            {
                get { return _key; }
                set { this._key = value; }
            }

            public int Verifier
            {
                get { return _verifier; }
                set { this._verifier = value; }
            }


            public void AppendToString(StringBuilder buffer)
            {
                buffer.Append("    .xor.key = ").Append(HexDump.IntToHex(_key)).Append("\n");
                buffer.Append("    .xor.verifier  = ").Append(HexDump.IntToHex(_verifier)).Append("\n");
            }

            /// <summary>
            /// 
            /// </summary>
            /// <returns>XorKeyData</returns>
            public object Clone()
            {
                XorKeyData other = new XorKeyData();
                other._key = this._key;
                other._verifier = this._verifier;
                return other;
            }
        }


        private FilePassRecord(FilePassRecord other)
        {
            _encryptionType = other._encryptionType;
            _keyData = (KeyData)other._keyData.Clone();
        }

        public FilePassRecord(RecordInputStream in1)
        {
            _encryptionType = in1.ReadUShort();

            switch (_encryptionType)
            {
                case ENCRYPTION_XOR:
                    _keyData = new XorKeyData();
                    break;
                case ENCRYPTION_OTHER:
                    _keyData = new Rc4KeyData();
                    break;
                default:
                    throw new RecordFormatException("Unknown encryption type " + _encryptionType);
            }

            _keyData.Read(in1);
        }

        private static byte[] Read(RecordInputStream in1, int size)
        {
            byte[] result = new byte[size];
            in1.ReadFully(result);
            return result;
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(_encryptionType);
            Debug.Assert(_keyData != null);
            _keyData.Serialize(out1);
        }

        protected override int DataSize
        {
            get
            {
                Debug.Assert(_keyData != null);
                return _keyData.DataSize;
            }
        }

        public Rc4KeyData GetRc4KeyData()
        {
            return (_keyData is Rc4KeyData)
                ? (Rc4KeyData)_keyData
                : null;
        }

        public XorKeyData GetXorKeyData()
        {
            return (_keyData is XorKeyData)
                ? (XorKeyData)_keyData
                : null;
        }

        private Rc4KeyData CheckRc4()
        {
            Rc4KeyData rc4 = GetRc4KeyData();
            if (rc4 == null)
            {
                throw new RecordFormatException("file pass record doesn't contain a rc4 key.");
            }
            return rc4;
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns>FilePassRecord</returns>
        public override object Clone()
        {
            return new FilePassRecord(this);
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FILEPASS]\n");
            buffer.Append("    .type = ").Append(HexDump.ShortToHex(_encryptionType)).Append("\n");
            _keyData.AppendToString(buffer);
            buffer.Append("[/FILEPASS]\n");
            return buffer.ToString();
        }
    }

}