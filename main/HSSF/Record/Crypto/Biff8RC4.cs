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

namespace NPOI.HSSF.Record.Crypto
{
    using System;
    using NPOI.HSSF.Record;

    /**
     * Used for both encrypting and decrypting BIFF8 streams. The internal
     * {@link RC4} instance is renewed (re-keyed) every 1024 bytes.
     *
     * @author Josh Micich
     */
    public class Biff8RC4
    {

        private const int RC4_REKEYING_INTERVAL = 1024;

        private RC4 _rc4;
        /**
         * This field is used to keep track of when to change the {@link RC4}
         * instance. The change occurs every 1024 bytes. Every byte passed over is
         * counted.
         */
        private int _streamPos;
        private int _nextRC4BlockStart;
        private int _currentKeyIndex;
        private bool _shouldSkipEncryptionOnCurrentRecord;

        private Biff8EncryptionKey _key;

        public Biff8RC4(int InitialOffset, Biff8EncryptionKey key)
        {
            if (InitialOffset >= RC4_REKEYING_INTERVAL)
            {
                throw new Exception("InitialOffset (" + InitialOffset + ")>"
                        + RC4_REKEYING_INTERVAL + " not supported yet");
            }
            _key = key;
            _streamPos = 0;
            RekeyForNextBlock();
            _streamPos = InitialOffset;
            for (int i = InitialOffset; i > 0; i--)
            {
                _rc4.Output();
            }
            _shouldSkipEncryptionOnCurrentRecord = false;
        }

        private void RekeyForNextBlock()
        {
            _currentKeyIndex = _streamPos / RC4_REKEYING_INTERVAL;
            _rc4 = _key.CreateRC4(_currentKeyIndex);
            _nextRC4BlockStart = (_currentKeyIndex + 1) * RC4_REKEYING_INTERVAL;
        }

        private int GetNextRC4Byte()
        {
            if (_streamPos >= _nextRC4BlockStart)
            {
                RekeyForNextBlock();
            }
            byte mask = _rc4.Output();
            _streamPos++;
            if (_shouldSkipEncryptionOnCurrentRecord)
            {
                return 0;
            }
            return mask & 0xFF;
        }

        public void StartRecord(int currentSid)
        {
            _shouldSkipEncryptionOnCurrentRecord = IsNeverEncryptedRecord(currentSid);
        }

        /**
         * TODO: Additionally, the lbPlyPos (position_of_BOF) field of the BoundSheet8 record MUST NOT be encrypted.
         *
         * @return <c>true</c> if record type specified by <c>sid</c> is never encrypted
         */
        private static bool IsNeverEncryptedRecord(int sid)
        {
            switch (sid)
            {
                case BOFRecord.sid:
                // sheet BOFs for sure
                // TODO - find out about chart BOFs

                case InterfaceHdrRecord.sid:
                // don't know why this record doesn't seem to get encrypted

                case FilePassRecord.sid:
                    // this only really counts when writing because FILEPASS is read early

                    // UsrExcl(0x0194)
                    // FileLock
                    // RRDInfo(0x0196)
                    // RRDHead(0x0138)

                    return true;
            }
            return false;
        }

        /**
         * Used when BIFF header fields (sid, size) are being Read. The internal
         * {@link RC4} instance must step even when unencrypted bytes are read
         */
        public void SkipTwoBytes()
        {
            GetNextRC4Byte();
            GetNextRC4Byte();
        }

        public void Xor(byte[] buf, int pOffSet, int pLen)
        {
            int nLeftInBlock;
            nLeftInBlock = _nextRC4BlockStart - _streamPos;
            if (pLen <= nLeftInBlock)
            {
                // simple case - this read does not cross key blocks
                _rc4.Encrypt(buf, pOffSet, pLen);
                _streamPos += pLen;
                return;
            }

            int offset = pOffSet;
            int len = pLen;

            // start by using the rest of the current block
            if (len > nLeftInBlock)
            {
                if (nLeftInBlock > 0)
                {
                    _rc4.Encrypt(buf, offset, nLeftInBlock);
                    _streamPos += nLeftInBlock;
                    offset += nLeftInBlock;
                    len -= nLeftInBlock;
                }
                RekeyForNextBlock();
            }
            // all full blocks following
            while (len > RC4_REKEYING_INTERVAL)
            {
                _rc4.Encrypt(buf, offset, RC4_REKEYING_INTERVAL);
                _streamPos += RC4_REKEYING_INTERVAL;
                offset += RC4_REKEYING_INTERVAL;
                len -= RC4_REKEYING_INTERVAL;
                RekeyForNextBlock();
            }
            // finish with incomplete block
            _rc4.Encrypt(buf, offset, len);
            _streamPos += len;
        }

        public int XorByte(int rawVal)
        {
            int mask = GetNextRC4Byte();
            return (byte)(rawVal ^ mask);
        }

        public int Xorshort(int rawVal)
        {
            int b0 = GetNextRC4Byte();
            int b1 = GetNextRC4Byte();
            int mask = (b1 << 8) + (b0 << 0);
            return rawVal ^ mask;
        }

        public int XorInt(int rawVal)
        {
            int b0 = GetNextRC4Byte();
            int b1 = GetNextRC4Byte();
            int b2 = GetNextRC4Byte();
            int b3 = GetNextRC4Byte();
            int mask = (b3 << 24) + (b2 << 16) + (b1 << 8) + (b0 << 0);
            return rawVal ^ mask;
        }

        public long XorLong(long rawVal)
        {
            int b0 = GetNextRC4Byte();
            int b1 = GetNextRC4Byte();
            int b2 = GetNextRC4Byte();
            int b3 = GetNextRC4Byte();
            int b4 = GetNextRC4Byte();
            int b5 = GetNextRC4Byte();
            int b6 = GetNextRC4Byte();
            int b7 = GetNextRC4Byte();
            long mask =
                  (((long)b7) << 56)
                + (((long)b6) << 48)
                + (((long)b5) << 40)
                + (((long)b4) << 32)
                + (((long)b3) << 24)
                + (b2 << 16)
                + (b1 << 8)
                + (b0 << 0);
            return rawVal ^ mask;
        }
    }
}

