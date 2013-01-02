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

namespace TestCases.Util
{
    using System;
    using System.IO;
    using NPOI.Util;


    using NUnit.Framework;
    /**
     * Class to test {@link LittleEndianInputStream} and {@link LittleEndianOutputStream}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestLittleEndianStreams
    {
        /// <summary>
        ///  Some of the tests are depending on the american culture.
        /// </summary>
        [SetUp]
        public void InitializeCultere()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }

        [Test]
        public void TestRead()
        {
            MemoryStream baos = new MemoryStream();
            ILittleEndianOutput leo = new LittleEndianOutputStream(baos);
            leo.WriteInt(12345678);
            leo.WriteShort(12345);
            leo.WriteByte(123);
            leo.WriteShort(40000);
            leo.WriteByte(200);
            leo.WriteLong(1234567890123456789L);
            leo.WriteDouble(123.456);

            ILittleEndianInput lei = new LittleEndianInputStream(new MemoryStream(baos.ToArray()));

            Assert.AreEqual(12345678, lei.ReadInt());
            Assert.AreEqual(12345, lei.ReadShort());
            Assert.AreEqual(123, lei.ReadByte());
            Assert.AreEqual(40000, lei.ReadUShort());
            Assert.AreEqual(200, lei.ReadUByte());
            Assert.AreEqual(1234567890123456789L, lei.ReadLong());
            Assert.AreEqual(123.456, lei.ReadDouble(), 0.0);
        }

        /**
         * Up until svn r836101 {@link LittleEndianMemoryStream#readFully(byte[], int, int)}
         * had an error which resulted in the data being read and written back to the source byte
         * array.
         */
        [Test]
        public void TestReadFully()
        {
            byte[] srcBuf = HexRead.ReadFromString("99 88 77 66 55 44 33");
            ILittleEndianInput lei = new LittleEndianByteArrayInputStream(srcBuf);

            // do Initial read to increment the read index beyond zero
            Assert.AreEqual(0x8899, lei.ReadUShort());

            byte[] actBuf = new byte[4];
            lei.ReadFully(actBuf);

            if (actBuf[0] == 0x00 && srcBuf[0] == 0x77 && srcBuf[3] == 0x44)
            {
                throw new AssertionException("Identified bug in ReadFully() - source buffer was modified");
            }

            byte[] expBuf = HexRead.ReadFromString("77 66 55 44");
            Assert.IsTrue(Arrays.Equals(actBuf, expBuf));
            Assert.AreEqual(0x33, lei.ReadUByte());
            Assert.AreEqual(0, lei.Available());
        }
    }
}
