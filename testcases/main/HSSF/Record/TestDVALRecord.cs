/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record
{
    using System.IO;
    using NPOI.Util;
    using NPOI.HSSF.Record;
    using NUnit.Framework;

    /**
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestDVALRecord
    {
        [Test]
        public void TestRead()
        {

            byte[] data = new byte[22];
            LittleEndian.PutShort(data, 0, DVALRecord.sid);
            LittleEndian.PutShort(data, 2, (short)18);
            LittleEndian.PutShort(data, 4, (short)55);
            LittleEndian.PutInt(data, 6, 56);
            LittleEndian.PutInt(data, 10, 57);
            LittleEndian.PutInt(data, 14, 58);
            LittleEndian.PutInt(data, 18, 59);

            RecordInputStream in1 = new RecordInputStream(new MemoryStream(data));
            in1.NextRecord();
            DVALRecord dv = new DVALRecord(in1);

            Assert.AreEqual(55, dv.Options);
            Assert.AreEqual(56, dv.HorizontalPos);
            Assert.AreEqual(57, dv.VerticalPos);
            Assert.AreEqual(58, dv.ObjectID);
            if (dv.DVRecNo == 0)
            {
                Assert.Fail("Identified bug 44510");
            }
            Assert.AreEqual(59, dv.DVRecNo);
        }
    }
}