/* ====================================================================
Licensed to the Apache Software Foundation (ASF) under one or more
contributor license agreements.  See the NOTICE file distributed with
this work for additional information regarding copyright ownership.
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
using NPOI.SS.Util;
using NUnit.Framework;
using TestCases.HSSF.Record;
using NPOI.Util;
using System.IO;
namespace TestCases.SS.Util
{
    //import java.io.ByteArrayOutputStream;

    //import org.apache.poi.hssf.record.TestcaseRecordInputStream;
    //import org.apache.poi.ss.util.CellRangeAddress;
    //import org.apache.poi.util.LittleEndianOutputStream;

    [TestFixture]
    public class TestCellRangeAddress
    {
        byte[] data = new byte[] {
     (byte)0x02,(byte)0x00, 
     (byte)0x04,(byte)0x00, 
     (byte)0x00,(byte)0x00, 
     (byte)0x03,(byte)0x00, 
 };
        [Test]
        public void TestLoad()
        {
            CellRangeAddress cref = new CellRangeAddress(
                 TestcaseRecordInputStream.Create(0x000, data)
           );
            Assert.AreEqual(2, cref.FirstRow);
            Assert.AreEqual(4, cref.LastRow);
            Assert.AreEqual(0, cref.FirstColumn);
            Assert.AreEqual(3, cref.LastColumn);

            Assert.AreEqual(8, CellRangeAddress.ENCODED_SIZE);
        }
        [Test]
        public void TestStore()
        {
            CellRangeAddress cref = new CellRangeAddress(0, 0, 0, 0);

            byte[] recordBytes;
            //ByteArrayOutputStream baos = new ByteArrayOutputStream();
            MemoryStream baos = new MemoryStream();
            LittleEndianOutputStream output = new LittleEndianOutputStream(baos);

            // With nothing set
            cref.Serialize(output);
            recordBytes = baos.ToArray();
            Assert.AreEqual(recordBytes.Length, data.Length);
            for (int i = 0; i < data.Length; i++)
            {
                Assert.AreEqual(0, recordBytes[i], "At offset " + i);
            }

            // Now set the flags
            cref.FirstRow = ((short)2);
            cref.LastRow = ((short)4);
            cref.FirstColumn = ((short)0);
            cref.LastColumn = ((short)3);

            // Re-test
            //baos.reset();
            baos.Seek(0, SeekOrigin.Begin);
            cref.Serialize(output);
            recordBytes = baos.ToArray();

            Assert.AreEqual(recordBytes.Length, data.Length);
            for (int i = 0; i < data.Length; i++)
            {
                Assert.AreEqual(data[i], recordBytes[i], "At offset " + i);
            }
        }
    }
}