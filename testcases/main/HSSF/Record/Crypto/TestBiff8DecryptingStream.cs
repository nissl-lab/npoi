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

namespace TestCases.HSSF.Record.Crypto
{
    using System;
    using System.IO;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NPOI.Util;
    using TestCases.Exceptions;
    using NPOI.HSSF.Record.Crypto;

    /**
     * Tests for {@link Biff8DecryptingStream}
     *
     * @author Josh Micich
     */
    [TestFixture]
    [Ignore("Ignored as RC4 encryption/decryption is validated in TestEncryptor.cs")]
    public class TestBiff8DecryptingStream
    {

        /**
         * A mock {@link InputStream} that keeps track of position and also produces
         * slightly interesting data. Each successive data byte value is one greater
         * than the previous.
         */
        private class MockStream : MemoryStream
        {
            private int _val;
            private int _position;

            public MockStream(int InitialValue)
            {
                _val = InitialValue & 0xFF;
            }
            public override int ReadByte()
            {
                _position++;
                return _val++ & 0xFF;
            }
            public int GetPosition()
            {
                    return _position;
            }
        }

        private class StreamTester
        {
            private static bool ONLY_LOG_ERRORS = true;

            private MockStream _ms;
            private Biff8DecryptingStream _bds;
            private bool _errorsOccurred;

            /**
             * @param expectedFirstInt expected value of the first int read from the decrypted stream
             */
            public StreamTester(MockStream ms, String keyDigestHex, int expectedFirstInt)
            {
                _ms = ms;
                byte[] keyDigest = HexRead.ReadFromString(keyDigestHex);
                //_bds = new Biff8DecryptingStream(_ms, 0, new Biff8EncryptionKey(keyDigest));
                ClassicAssert.AreEqual(expectedFirstInt, _bds.ReadInt());
                _errorsOccurred = false;
            }

            public Biff8DecryptingStream GetBDS()
            {
                return _bds;
            }

            /**
             * Used to 'skip over' the uninteresting middle bits of the key blocks.
             * Also Confirms that read position of the underlying stream is aligned.
             */
            public void RollForward(int fromPosition, int toPosition)
            {
                ClassicAssert.AreEqual(fromPosition, _ms.GetPosition());
                for (int i = fromPosition; i < toPosition; i++)
                {
                    _bds.ReadByte();
                }
                ClassicAssert.AreEqual(toPosition, _ms.GetPosition());
            }

            public void ConfirmByte(int expVal)
            {
                cmp(HexDump.ByteToHex(expVal), HexDump.ByteToHex(_bds.ReadUByte()));
            }

            public void ConfirmShort(int expVal)
            {
                cmp(HexDump.ShortToHex(expVal), HexDump.ShortToHex(_bds.ReadUShort()));
            }

            public void ConfirmUShort(int expVal)
            {
                cmp(HexDump.ShortToHex(expVal), HexDump.ShortToHex(_bds.ReadUShort()));
            }

            public short ReadShort()
            {
                return _bds.ReadShort();
            }

            public int ReadUShort()
            {
                return _bds.ReadUShort();
            }

            public void ConfirmInt(int expVal)
            {
                cmp(HexDump.IntToHex(expVal), HexDump.IntToHex(_bds.ReadInt()));
            }

            public void ConfirmLong(long expVal)
            {
                cmp(HexDump.LongToHex(expVal), HexDump.LongToHex(_bds.ReadLong()));
            }

            private void cmp(char[] exp, char[] act)
            {
                if (Arrays.Equals(exp, act))
                {
                    return;
                }
                _errorsOccurred = true;
                if (ONLY_LOG_ERRORS)
                {
                    logErr(3, "Value mismatch " + new String(exp) + " - " + new String(act));
                    return;
                }
                throw new ComparisonFailure("Value mismatch", new String(exp), new String(act));
            }

            public void ConfirmData(String expHexData)
            {

                byte[] expData = HexRead.ReadFromString(expHexData);
                byte[] actData = new byte[expData.Length];
                _bds.ReadFully(actData);
                if (Arrays.Equals(expData, actData))
                {
                    return;
                }
                _errorsOccurred = true;
                if (ONLY_LOG_ERRORS)
                {
                    logErr(2, "Data mismatch " + HexDump.ToHex(expData) + " - "
                            + HexDump.ToHex(actData));
                    return;
                }
                throw new ComparisonFailure("Data mismatch", HexDump.ToHex(expData), HexDump.ToHex(actData));
            }

            private static void logErr(int stackFrameCount, String msg)
            {
               // StackTraceElement ste = new Exception().StackTrace[stackFrameCount];
                //System.err.Print("(" + ste.FileName + ":" + ste.LineNumber + ") ");
                //System.err.Println(msg);
            }

            public void AssertNoErrors()
            {
                ClassicAssert.IsFalse(_errorsOccurred, "Some values decrypted incorrectly");
            }
        }

        /**
         * Tests Reading of 64,32,16 and 8 bit integers aligned with key changing boundaries
         */
        [Test]
        [Ignore("Ignored as RC4 encryption/decryption is validated in TestEncryptor.cs")]
        public void TestReadsAlignedWithBoundary()
        {
            StreamTester st = CreateStreamTester(0x50, "BA AD F0 0D 00",  unchecked((int)0x96C66829));

            st.RollForward(0x0004, 0x03FF);
            st.ConfirmByte(0x3E);
            st.ConfirmByte(0x28);
            st.RollForward(0x0401, 0x07FE);
            st.ConfirmShort(0x76CC);
            st.ConfirmShort(0xD83E);
            st.RollForward(0x0802, 0x0BFC);
            st.ConfirmInt(0x25F280EB);
            st.ConfirmInt(unchecked((int)0xB549E99B));
            st.RollForward(0x0C04, 0x0FF8);
            st.ConfirmLong(0x6AA2D5F6B975D10CL);
            st.ConfirmLong(0x34248ADF7ED4F029L);

            // check for signed/unsigned shorts #58069
            st.RollForward(0x1008, 0x7213);
            st.ConfirmUShort(0xFFFF);
            st.RollForward(0x7215, 0x1B9AD);
            st.ConfirmShort(-1);
            st.RollForward(0x1B9AF, 0x37D99);
            ClassicAssert.AreEqual(0xFFFF, st.ReadUShort());
            st.RollForward(0x37D9B, 0x4A6F2);
            ClassicAssert.AreEqual(-1, st.ReadShort());

            st.AssertNoErrors();
        }

        /**
         * Tests Reading of 64,32 and 16 bit integers <i>across</i> key changing boundaries
         */
        [Test]
        [Ignore("Ignored as RC4 encryption/decryption is validated in TestEncryptor.cs")]
        public void TestReadsSpanningBoundary()
        {
            StreamTester st = CreateStreamTester(0x50, "BA AD F0 0D 00",  unchecked((int)0x96C66829));

            st.RollForward(0x0004, 0x03FC);
            st.ConfirmLong(unchecked((long)0x885243283E2A5EEFL));
            st.RollForward(0x0404, 0x07FE);
            st.ConfirmInt( unchecked((int)0xD83E76CC));
            st.RollForward(0x0802, 0x0BFF);
            st.ConfirmShort(0x9B25);
            st.AssertNoErrors();
        }

        /**
         * Checks that the BIFF header fields (sid, size) Get read without Applying decryption,
         * and that the RC4 stream stays aligned during these calls
         */
        [Test]
        [Ignore("Ignored as RC4 encryption/decryption is validated in TestEncryptor.cs")]
        public void TestReadHeaderUshort()
        {
            StreamTester st = CreateStreamTester(0x50, "BA AD F0 0D 00",  unchecked((int)0x96C66829));

            st.RollForward(0x0004, 0x03FF);

            Biff8DecryptingStream bds = st.GetBDS();
            int hval = bds.ReadDataSize();   // unencrypted
            int nextInt = bds.ReadInt();
            if (nextInt == unchecked((int)0x8F534029))
            {
                throw new AssertionException(
                        "Indentified bug in key alignment After call to ReadHeaderUshort()");
            }
            ClassicAssert.AreEqual(0x16885243, nextInt);
            if (hval == 0x283E)
            {
                throw new AssertionException("readHeaderUshort() incorrectly decrypted result");
            }
            ClassicAssert.AreEqual(0x504F, hval);

            // confirm next key change
            st.RollForward(0x0405, 0x07FC);
            st.ConfirmInt(0x76CC1223);
            st.ConfirmInt(0x4842D83E);
            st.AssertNoErrors();
        }

        /**
         * Tests Reading of byte sequences <i>across</i> and <i>aligned with</i> key changing boundaries
         */
        [Test]
        [Ignore("Ignored as RC4 encryption/decryption is validated in TestEncryptor.cs")]
        public void TestReadByteArrays()
        {
            StreamTester st = CreateStreamTester(0x50, "BA AD F0 0D 00", unchecked((int)0x96C66829));

            st.RollForward(0x0004, 0x2FFC);
            st.ConfirmData("66 A1 20 B1 04 A3 35 F5"); // 4 bytes on either side of boundary
            st.RollForward(0x3004, 0x33F8);
            st.ConfirmData("F8 97 59 36");  // last 4 bytes in block
            st.ConfirmData("01 C2 4E 55");  // first 4 bytes in next block
            st.AssertNoErrors();
        }

        private static StreamTester CreateStreamTester(int mockStreamStartVal, String keyDigestHex, int expectedFirstInt)
        {
            return new StreamTester(new MockStream(mockStreamStartVal), keyDigestHex, expectedFirstInt);
        }
    }

}