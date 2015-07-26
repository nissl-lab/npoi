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

namespace TestCases.HSSF.Record
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using NUnit.Framework;
    using NPOI.Util;
    using NPOI.HSSF.Record;



    /**
     * Tests the serialization and deserialization of the EndSubRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestInterfaceEndRecord
    {
        [Test]
        public void TestCreate()
        {
            InterfaceEndRecord record = InterfaceEndRecord.Instance;
            Assert.AreEqual(0, record.GetDataSize());
        }

        /**
         * Silently swallow unexpected contents in InterfaceEndRecord.
         * Although it violates the spec, Excel silently Converts this
         * data to an {@link InterfaceHdrRecord}.
         */
        [Test]
        public void TestUnexpectedBytes_bug47251()
        {
            String hex = "" +
                    "09 08 10 00 00 06 05 00 EC 15 CD 07 C1 C0 00 00 06 03 00 00 " +   //BOF
                    "E2 00 02 00 B0 04 " + //INTERFACEEND with extra two bytes
                    "0A 00 00 00";    // EOF
            byte[] data = HexRead.ReadFromString(hex);
            List<Record> records = RecordFactory.CreateRecords(new MemoryStream(data));
            Assert.AreEqual(3, records.Count);
            Record rec1 = records[(1)];
            Assert.AreEqual(typeof(InterfaceHdrRecord), rec1.GetType());
            InterfaceHdrRecord r = (InterfaceHdrRecord)rec1;
            Assert.AreEqual("[E1, 00, 02, 00, B0, 04]", HexDump.ToHex(r.Serialize()));
        }
    }

}