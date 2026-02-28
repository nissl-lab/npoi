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
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NPOI.HSSF.Record;
    using NPOI.Util;

    [TestFixture]
    public class TestDrawingRecord
    {

        /**
         * Check that RecordFactoryInputStream properly handles continued DrawingRecords
         * See Bugzilla #47548
         */
        [Test]
        public void TestReadContinued()
        {

            //simulate a continues Drawing record
            MemoryStream out1 = new MemoryStream();
            //main part
            DrawingRecord dg = new DrawingRecord();
            byte[] data1 = new byte[8224];
            Arrays.Fill(data1, (byte)1);
            dg.SetData(data1);
            byte[] dataX = dg.Serialize();
            out1.Write(dataX, 0, dataX.Length);

            //continued part
            byte[] data2 = new byte[4048];
            Arrays.Fill(data2, (byte)2);
            ContinueRecord cn = new ContinueRecord(data2);
            dataX = cn.Serialize();
            out1.Write(dataX, 0, dataX.Length);

            List<Record> rec = RecordFactory.CreateRecords(new ByteArrayInputStream(out1.ToArray()));
            ClassicAssert.AreEqual(2, rec.Count);
            ClassicAssert.IsTrue(rec[0] is DrawingRecord);
            ClassicAssert.IsTrue(rec[1] is ContinueRecord);

            ClassicAssert.IsTrue(Arrays.Equals(data1, ((DrawingRecord)rec[0]).RecordData));
            ClassicAssert.IsTrue(Arrays.Equals(data2, ((ContinueRecord)rec[1]).Data));

        }

    }

}