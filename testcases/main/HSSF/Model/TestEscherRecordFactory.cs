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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NPOI.HSSF.Record;
using NPOI.Util;
using NUnit.Framework;
using NPOI.DDF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Model;
using TestCases.HSSF.UserModel;

namespace TestCases.HSSF.Model
{
    /**
 * @author Evgeniy Berlog
 * @date 18.06.12
 */
    [TestFixture]
    public class TestEscherRecordFactory
    {
        private static byte[] toByteArray(List<RecordBase> records)
        {
            MemoryStream out1 = new MemoryStream();
            foreach (RecordBase rb in records)
            {
                NPOI.HSSF.Record.Record r = (NPOI.HSSF.Record.Record)rb;
                try
                {
                    byte[] data = r.Serialize();
                    out1.Write(data, 0, data.Length);
                }
                catch (IOException e)
                {
                    throw new RuntimeException(e);
                }
            }
            return out1.ToArray();
        }
        [Test]
        public void TestDetectContainer()
        {
            Random rnd = new Random();
            Assert.AreEqual(true, DefaultEscherRecordFactory.IsContainer((short)0x0, EscherContainerRecord.DG_CONTAINER));
            Assert.AreEqual(true, DefaultEscherRecordFactory.IsContainer((short)0x0, EscherContainerRecord.SOLVER_CONTAINER));
            Assert.AreEqual(true, DefaultEscherRecordFactory.IsContainer((short)0x0, EscherContainerRecord.SP_CONTAINER));
            Assert.AreEqual(true, DefaultEscherRecordFactory.IsContainer((short)0x0, EscherContainerRecord.DGG_CONTAINER));
            Assert.AreEqual(true, DefaultEscherRecordFactory.IsContainer((short)0x0, EscherContainerRecord.BSTORE_CONTAINER));
            Assert.AreEqual(true, DefaultEscherRecordFactory.IsContainer((short)0x0, EscherContainerRecord.SPGR_CONTAINER));

            for (short i = EscherContainerRecord.DGG_CONTAINER; i <= EscherContainerRecord.SOLVER_CONTAINER; i++)
            {
                Assert.AreEqual(true, DefaultEscherRecordFactory.IsContainer((short)rnd.Next(short.MaxValue), i));
            }

            Assert.AreEqual(false, DefaultEscherRecordFactory.IsContainer((short)0x0, EscherContainerRecord.DGG_CONTAINER - 1));
            Assert.AreEqual(false, DefaultEscherRecordFactory.IsContainer((short)0x0, EscherContainerRecord.SOLVER_CONTAINER + 1));

            Assert.AreEqual(true, DefaultEscherRecordFactory.IsContainer((short)0x000F, EscherContainerRecord.DGG_CONTAINER - 1));
            Assert.AreEqual(true, DefaultEscherRecordFactory.IsContainer(unchecked((short)0xFFFF), EscherContainerRecord.DGG_CONTAINER - 1));
            Assert.AreEqual(false, DefaultEscherRecordFactory.IsContainer((short)0x000C, EscherContainerRecord.DGG_CONTAINER - 1));
            Assert.AreEqual(false, DefaultEscherRecordFactory.IsContainer(unchecked((short)0xCCCC), EscherContainerRecord.DGG_CONTAINER - 1));
            Assert.AreEqual(false, DefaultEscherRecordFactory.IsContainer((short)0x000F, EscherTextboxRecord.RECORD_ID));
            Assert.AreEqual(false, DefaultEscherRecordFactory.IsContainer(unchecked((short)0xCCCC), EscherTextboxRecord.RECORD_ID));
        }
        [Test]
        public void TestDgContainerMustBeRootOfHSSFSheetEscherRecords()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("47251.xls");
            HSSFSheet sh = wb.GetSheetAt(0) as HSSFSheet;
            InternalSheet ish = HSSFTestHelper.GetSheetForTest(sh);
            List<RecordBase> records = ish.Records;
            // records to be aggregated
            List<RecordBase> dgRecords = records.GetRange(19, 23 - 19);
            byte[] dgBytes = toByteArray(dgRecords);
            IDrawing d = sh.DrawingPatriarch;
            EscherAggregate agg = (EscherAggregate)ish.FindFirstRecordBySid(EscherAggregate.sid);
            Assert.AreEqual(true, agg.EscherRecords[0] is EscherContainerRecord);
            Assert.AreEqual(EscherContainerRecord.DG_CONTAINER, agg.EscherRecords[0].RecordId);
            Assert.AreEqual((short)0x0, agg.EscherRecords[0].Options);
            agg = (EscherAggregate)ish.FindFirstRecordBySid(EscherAggregate.sid);
            byte[] dgBytesAfterSave = agg.Serialize();
            Assert.AreEqual(dgBytes.Length, dgBytesAfterSave.Length, "different size of drawing data before and after save");
            Assert.IsTrue(Arrays.Equals(dgBytes, dgBytesAfterSave), "drawing data before and after save is different");
        }
    }
}
