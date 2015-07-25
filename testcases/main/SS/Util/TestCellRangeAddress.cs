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
using System;
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
        public void TestLoadInvalid()
        {
            try
            {
                Assert.IsNotNull(new CellRangeAddress(
                    TestcaseRecordInputStream.Create(0x000, new byte[] { (byte)0x02 })));
            }
            catch (RuntimeException e)
            {
                Assert.IsTrue(e.Message.Contains("Ran out of data"), "Had: " + e);
            }
        }
        [Test]
        public void TestStore()
        {
            CellRangeAddress cref = new CellRangeAddress(0, 0, 0, 0);

            byte[] recordBytes;
            //ByteArrayOutputStream baos = new ByteArrayOutputStream();
            MemoryStream baos = new MemoryStream();
            LittleEndianOutputStream output = new LittleEndianOutputStream(baos);
            try
            {
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
            finally
            {
                //output.Close();
            }
        }

        [Test]
        public void TestStoreDeprecated()
        {
            CellRangeAddress ref1 = new CellRangeAddress(0, 0, 0, 0);

            //byte[] recordBytes = new byte[CellRangeAddress.ENCODED_SIZE];
            //// With nothing Set
            //ref1.Serialize(0, recordBytes);
            //Assert.AreEqual(recordBytes.Length, data.Length);
            //for (int i = 0; i < data.Length; i++)
            //{
            //    Assert.AreEqual("At offset " + i, 0, recordBytes[i]);
            //}

            //// Now Set the flags
            //ref1.FirstRow = (/*setter*/(short)2);
            //ref1.LastRow = (/*setter*/(short)4);
            //ref1.FirstColumn = (/*setter*/(short)0);
            //ref1.LastColumn = (/*setter*/(short)3);

            //// Re-test
            //ref1.Serialize(0, recordBytes);

            //Assert.AreEqual(recordBytes.Length, data.Length);
            //for (int i = 0; i < data.Length; i++)
            //{
            //    Assert.AreEqual("At offset " + i, data[i], recordBytes[i]);
            //}
        }

        [Test]
        public void TestCreateIllegal()
        {
            // for some combinations we expected exceptions
            try
            {
                Assert.IsNotNull(new CellRangeAddress(1, 0, 0, 0));
                Assert.Fail("Expect to catch an exception");
            }
            catch (ArgumentException)
            {
                // expected here
            }
            try
            {
                Assert.IsNotNull(new CellRangeAddress(0, 0, 1, 0));
                Assert.Fail("Expect to catch an exception");
            }
            catch (ArgumentException)
            {
                // expected here
            }
        }

        [Test]
        public void TestCopy()
        {
            CellRangeAddress ref1 = new CellRangeAddress(1, 2, 3, 4);
            CellRangeAddress copy = ref1.Copy();
            Assert.AreEqual(ref1.ToString(), copy.ToString());
        }

        [Test]
        public void TestGetEncodedSize()
        {
            Assert.AreEqual(2 * CellRangeAddress.ENCODED_SIZE, CellRangeAddress.GetEncodedSize(2));
        }

        [Test]
        public void TestFormatAsString()
        {
            CellRangeAddress ref1 = new CellRangeAddress(1, 2, 3, 4);

            Assert.AreEqual("D2:E3", ref1.FormatAsString());
            Assert.AreEqual("D2:E3", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString());

            Assert.AreEqual("sheet1!$D$2:$E$3", ref1.FormatAsString("sheet1", true));
            Assert.AreEqual("sheet1!$D$2:$E$3", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString("sheet1", true));
            Assert.AreEqual("sheet1!$D$2:$E$3", CellRangeAddress.ValueOf(ref1.FormatAsString("sheet1", true)).FormatAsString("sheet1", true));

            Assert.AreEqual("sheet1!D2:E3", ref1.FormatAsString("sheet1", false));
            Assert.AreEqual("sheet1!D2:E3", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString("sheet1", false));
            Assert.AreEqual("sheet1!D2:E3", CellRangeAddress.ValueOf(ref1.FormatAsString("sheet1", false)).FormatAsString("sheet1", false));

            Assert.AreEqual("D2:E3", ref1.FormatAsString(null, false));
            Assert.AreEqual("D2:E3", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString(null, false));
            Assert.AreEqual("D2:E3", CellRangeAddress.ValueOf(ref1.FormatAsString(null, false)).FormatAsString(null, false));

            Assert.AreEqual("$D$2:$E$3", ref1.FormatAsString(null, true));
            Assert.AreEqual("$D$2:$E$3", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString(null, true));
            Assert.AreEqual("$D$2:$E$3", CellRangeAddress.ValueOf(ref1.FormatAsString(null, true)).FormatAsString(null, true));

            ref1 = new CellRangeAddress(-1, -1, 3, 4);
            Assert.AreEqual("D:E", ref1.FormatAsString());
            Assert.AreEqual("sheet1!$D:$E", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString("sheet1", true));
            Assert.AreEqual("sheet1!$D:$E", CellRangeAddress.ValueOf(ref1.FormatAsString("sheet1", true)).FormatAsString("sheet1", true));
            Assert.AreEqual("$D:$E", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString(null, true));
            Assert.AreEqual("$D:$E", CellRangeAddress.ValueOf(ref1.FormatAsString(null, true)).FormatAsString(null, true));
            Assert.AreEqual("sheet1!D:E", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString("sheet1", false));
            Assert.AreEqual("sheet1!D:E", CellRangeAddress.ValueOf(ref1.FormatAsString("sheet1", false)).FormatAsString("sheet1", false));

            ref1 = new CellRangeAddress(1, 2, -1, -1);
            Assert.AreEqual("2:3", ref1.FormatAsString());
            Assert.AreEqual("sheet1!$2:$3", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString("sheet1", true));
            Assert.AreEqual("sheet1!$2:$3", CellRangeAddress.ValueOf(ref1.FormatAsString("sheet1", true)).FormatAsString("sheet1", true));
            Assert.AreEqual("$2:$3", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString(null, true));
            Assert.AreEqual("$2:$3", CellRangeAddress.ValueOf(ref1.FormatAsString(null, true)).FormatAsString(null, true));
            Assert.AreEqual("sheet1!2:3", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString("sheet1", false));
            Assert.AreEqual("sheet1!2:3", CellRangeAddress.ValueOf(ref1.FormatAsString("sheet1", false)).FormatAsString("sheet1", false));

            ref1 = new CellRangeAddress(1, 1, 2, 2);
            Assert.AreEqual("C2", ref1.FormatAsString());
            Assert.AreEqual("sheet1!$C$2", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString("sheet1", true));
            Assert.AreEqual("sheet1!$C$2", CellRangeAddress.ValueOf(ref1.FormatAsString("sheet1", true)).FormatAsString("sheet1", true));
            Assert.AreEqual("$C$2", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString(null, true));
            Assert.AreEqual("$C$2", CellRangeAddress.ValueOf(ref1.FormatAsString(null, true)).FormatAsString(null, true));
            Assert.AreEqual("sheet1!C2", CellRangeAddress.ValueOf(ref1.FormatAsString()).FormatAsString("sheet1", false));
            Assert.AreEqual("sheet1!C2", CellRangeAddress.ValueOf(ref1.FormatAsString("sheet1", false)).FormatAsString("sheet1", false));

            // is this a valid Address?
            ref1 = new CellRangeAddress(-1, -1, -1, -1);
            Assert.AreEqual(":", ref1.FormatAsString());
        }

    }
}