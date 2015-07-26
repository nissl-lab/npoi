
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

    using NUnit.Framework;
    using NPOI.DDF;
    using NPOI.Util;

using System.Configuration;

    [TestFixture]
    public class TestEscherContainerRecord
    {
        private String ESCHER_DATA_PATH;

        public TestEscherContainerRecord()
        {
            ESCHER_DATA_PATH = ConfigurationManager.AppSettings["escher.data.path"];
        }
        [Test]
        public void TestFillFields()
        {
            IEscherRecordFactory f = new DefaultEscherRecordFactory();
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
        [Test]
        public void TestSerialize()
        {
            UnknownEscherRecord r = new UnknownEscherRecord();
            r.Options=(short)0x123F;
            r.RecordId=unchecked((short)0xF112);
            byte[] data = new byte[8];
            r.Serialize(0, data);

            Assert.AreEqual("[3F, 12, 12, F1, 00, 00, 00, 00]", HexDump.ToHex(data));

            EscherRecord childRecord = new UnknownEscherRecord();
            childRecord.Options=unchecked((short)0x9999);
            childRecord.RecordId=unchecked((short)0xFF01);
            r.AddChildRecord(childRecord);
            data = new byte[16];
            r.Serialize(0, data);

            Assert.AreEqual("[3F, 12, 12, F1, 08, 00, 00, 00, 99, 99, 01, FF, 00, 00, 00, 00]", HexDump.ToHex(data));

        }
        [Test]
        public void TestToString()
        {
            EscherContainerRecord r = new EscherContainerRecord();
            r.RecordId=EscherContainerRecord.SP_CONTAINER;
            r.Options=(short)0x000F;
            String nl = Environment.NewLine;
            Assert.AreEqual("EscherContainerRecord (SpContainer):" + nl +
                    "  isContainer: True" + nl +
                    "  version: 0x000F" + nl +
                    "  instance: 0x0000" + nl +
                    "  recordId: 0xF004" + nl +
                    "  numchildren: 0" + nl
                    , r.ToString());

            EscherOptRecord r2 = new EscherOptRecord();
            // don't try to shoot in foot, please -- vlsergey
            // r2.setOptions((short) 0x9876);
            r2.RecordId=EscherOptRecord.RECORD_ID;

            String expected;
            r.AddChildRecord(r2);
            expected = "EscherContainerRecord (SpContainer):" + nl +
                       "  isContainer: True" + nl +
                       "  version: 0x000F" + nl +
                       "  instance: 0x0000" + nl +
                       "  recordId: 0xF004" + nl +
                       "  numchildren: 1" + nl +
                       "  children: " + nl +
                       "    Child 0:" + nl +
                       "    EscherOptRecord:" + nl +
                       "      isContainer: False" + nl +
                       "      version: 0x0003" + nl +
                       "      instance: 0x0000" + nl +
                       "      recordId: 0xF00B" + nl +
                       "      numchildren: 0" + nl +
                       "      properties:" + nl +
                       "    " + nl;
            Assert.AreEqual(expected, r.ToString());

            r.AddChildRecord(r2);
            expected = "EscherContainerRecord (SpContainer):" + nl +
                       "  isContainer: True" + nl +
                       "  version: 0x000F" + nl +
                       "  instance: 0x0000" + nl +
                       "  recordId: 0xF004" + nl +
                       "  numchildren: 2" + nl +
                       "  children: " + nl +
                       "    Child 0:" + nl +
                       "    EscherOptRecord:" + nl +
                       "      isContainer: False" + nl +
                       "      version: 0x0003" + nl +
                       "      instance: 0x0000" + nl +
                       "      recordId: 0xF00B" + nl +
                       "      numchildren: 0" + nl +
                       "      properties:" + nl +
                       "    " + nl +
                       "    Child 1:" + nl +
                       "    EscherOptRecord:" + nl +
                       "      isContainer: False" + nl +
                       "      version: 0x0003" + nl +
                       "      instance: 0x0000" + nl +
                       "      recordId: 0xF00B" + nl +
                       "      numchildren: 0" + nl +
                       "      properties:" + nl +
                       "    " + nl;
            Assert.AreEqual(expected, r.ToString());
        }

        private class DummyEscherRecord : EscherRecord
        {
            public DummyEscherRecord() { }
            public override int FillFields(byte[] data, int offset, IEscherRecordFactory recordFactory)
            { return 0; }
            public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
            { return 0; }
            public override int RecordSize { get { return 10; } }
            public override String RecordName { get { return ""; } }
        }
        [Test]
        public void TestRecordSize()
        {
            EscherContainerRecord r = new EscherContainerRecord();

            r.AddChildRecord(new DummyEscherRecord());

            Assert.AreEqual(18, r.RecordSize);
        }

        /**
         * We were having problems with Reading too much data on an UnknownEscherRecord,
         *  but hopefully we now Read the correct size.
         */
        [Test]
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
        /**
	 * Ensure {@link EscherContainerRecord} doesn't spill its guts everywhere
	 */
        [Test]
        public void TestChildren()
        {
            EscherContainerRecord ecr = new EscherContainerRecord();
            List<EscherRecord> children0 = ecr.ChildRecords;
            Assert.AreEqual(0, children0.Count);

            EscherRecord chA = new DummyEscherRecord();
            EscherRecord chB = new DummyEscherRecord();
            EscherRecord chC = new DummyEscherRecord();

            ecr.AddChildRecord(chA);
            ecr.AddChildRecord(chB);
            children0.Add(chC);

            List<EscherRecord> children1 = ecr.ChildRecords;
            Assert.IsTrue(children0 != children1);
            Assert.AreEqual(2, children1.Count);
            Assert.AreEqual(chA, children1[0]);
            Assert.AreEqual(chB, children1[1]);

            Assert.AreEqual(1, children0.Count); // first copy unchanged

            ecr.ChildRecords=(children0);
            ecr.AddChildRecord(chA);
            List<EscherRecord> children2 = ecr.ChildRecords;
            Assert.AreEqual(2, children2.Count);
            Assert.AreEqual(chC, children2[0]);
            Assert.AreEqual(chA, children2[1]);
        }
    }
}