
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

namespace TestCases.DDF
{

    using System;
    using System.Text;
    using System.Collections.Generic;
    using System.IO;

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.DDF;
    using NPOI.Util;
    using NPOI.Util.IO;
using System.Configuration;

    [TestClass]
    public class TestEscherContainerRecord
    {
        private String ESCHER_DATA_PATH;

        public TestEscherContainerRecord()
        {
            ESCHER_DATA_PATH = ConfigurationManager.AppSettings["escher.data.path"];
        }
        [TestMethod]
        public void TestFillFields()
        {
            EscherRecordFactory f = new DefaultEscherRecordFactory();
            byte[] data = HexRead.ReadFromString("0F 02 11 F1 00 00 00 00");
            EscherRecord r = f.CreateRecord(data, 0);
            r.FillFields(data, 0, f);
            Assert.IsTrue(r is EscherContainerRecord);
            Assert.AreEqual((short)0x020F, r.Options);
            Assert.AreEqual(unchecked((short)0xF111), r.RecordId);

            data = HexRead.ReadFromString("0F 02 11 F1 08 00 00 00" +
                    " 02 00 22 F2 00 00 00 00");
            r = f.CreateRecord(data, 0);
            r.FillFields(data, 0, f);
            EscherRecord c = r.GetChild(0);
            Assert.IsFalse(c is EscherContainerRecord);
            Assert.AreEqual(unchecked((short)0x0002), c.Options);
            Assert.AreEqual(unchecked((short)0xF222), c.RecordId);
        }
        [TestMethod]
        public void TestSerialize()
        {
            UnknownEscherRecord r = new UnknownEscherRecord();
            r.Options=(short)0x123F;
            r.RecordId=unchecked((short)0xF112);
            byte[] data = new byte[8];
            r.Serialize(0, data);

            Assert.AreEqual("[3F, 12, 12, F1, 00, 00, 00, 00, ]", HexDump.ToHex(data));

            EscherRecord childRecord = new UnknownEscherRecord();
            childRecord.Options=unchecked((short)0x9999);
            childRecord.RecordId=unchecked((short)0xFF01);
            r.AddChildRecord(childRecord);
            data = new byte[16];
            r.Serialize(0, data);

            Assert.AreEqual("[3F, 12, 12, F1, 08, 00, 00, 00, 99, 99, 01, FF, 00, 00, 00, 00, ]", HexDump.ToHex(data));

        }
        [TestMethod]
        public void TestToString()
        {
            EscherContainerRecord r = new EscherContainerRecord();
            r.RecordId=EscherContainerRecord.SP_CONTAINER;
            r.Options=(short)0x000F;
            String nl = Environment.NewLine;
            Assert.AreEqual("EscherContainerRecord (SpContainer):" + nl +
                    "  isContainer: True" + nl +
                    "  options: 0x000F" + nl +
                    "  recordId: 0xF004" + nl +
                    "  numchildren: 0" + nl
                    , r.ToString());

            EscherOptRecord r2 = new EscherOptRecord();
            r2.Options = unchecked((short)0x9876) ;
            r2.RecordId=EscherOptRecord.RECORD_ID;

            String expected;
            r.AddChildRecord(r2);
            expected = "EscherContainerRecord (SpContainer):" + nl +
                       "  isContainer: True" + nl +
                       "  options: 0x000F" + nl +
                       "  recordId: 0xF004" + nl +
                       "  numchildren: 1" + nl +
                       "  children: " + nl +
                       "   Child 0:" + nl +
                       "EscherOptRecord:" + nl +
                       "  isContainer: False" + nl +
                       "  options: 0x0003" + nl +
                       "  recordId: 0xF00B" + nl +
                       "  numchildren: 0" + nl +
                       "  properties:" + nl;
            Assert.AreEqual(expected, r.ToString());

            r.AddChildRecord(r2);
            expected = "EscherContainerRecord (SpContainer):" + nl +
                       "  isContainer: True" + nl +
                       "  options: 0x000F" + nl +
                       "  recordId: 0xF004" + nl +
                       "  numchildren: 2" + nl +
                       "  children: " + nl +
                       "   Child 0:" + nl +
                       "EscherOptRecord:" + nl +
                       "  isContainer: False" + nl +
                       "  options: 0x0003" + nl +
                       "  recordId: 0xF00B" + nl +
                       "  numchildren: 0" + nl +
                       "  properties:" + nl +
                       "   Child 1:" + nl +
                       "EscherOptRecord:" + nl +
                       "  isContainer: False" + nl +
                       "  options: 0x0003" + nl +
                       "  recordId: 0xF00B" + nl +
                       "  numchildren: 0" + nl +
                       "  properties:" + nl;
            Assert.AreEqual(expected, r.ToString());
        }

        private class TestRecordA : EscherRecord
        {
            public override int FillFields(byte[] data, int offset, EscherRecordFactory recordFactory)
            { return 0; }
            public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
            { return 0; }
            public override int RecordSize { get { return 10; } }
            public override String RecordName { get { return ""; } }
        }
        [TestMethod]
        public void TestRecordSize()
        {
            EscherContainerRecord r = new EscherContainerRecord();

            r.AddChildRecord(new TestRecordA());

            Assert.AreEqual(18, r.RecordSize);
        }

        /**
         * We were having problems with Reading too much data on an UnknownEscherRecord,
         *  but hopefully we now Read the correct size.
         */
        [TestMethod]
        public void TestBug44857()
        {
            //File f = new File(ESCHER_DATA_PATH, "Container.dat");
            Assert.IsTrue(File.Exists(ESCHER_DATA_PATH+"Container.dat"));

            using (FileStream finp = new FileStream(ESCHER_DATA_PATH + "Container.dat", FileMode.Open, FileAccess.Read))
            {
                byte[] data = IOUtils.ToByteArray(finp);
                finp.Close();

                // This used to fail with an OutOfMemory
                EscherContainerRecord record = new EscherContainerRecord();
                record.FillFields(data, 0, new DefaultEscherRecordFactory());
            }
        }
    }
}