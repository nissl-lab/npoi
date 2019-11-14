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
        [Test]
        public void TestEquals()
        {
            CellRangeAddress ref1 = new CellRangeAddress(1, 2, 3, 4);
            CellRangeAddress ref2 = new CellRangeAddress(1, 2, 3, 4);
            Assert.AreEqual(ref1, ref2);

            // Invert first/last row, but refer to same area
            ref2.FirstRow = (2);
            ref2.LastRow = (1);
            Assert.AreEqual(ref1, ref2);

            // Invert first/last column, but refer to same area
            ref2.FirstColumn = (4);
            ref2.LastColumn = (3);
            Assert.AreEqual(ref1, ref2);

            // Refer to a different area
            Assert.AreNotEqual(ref1, new CellRangeAddress(3, 4, 1, 2));
        }
        [Test]
        public void TestGetMinMaxRow()
        {
            CellRangeAddress ref1 = new CellRangeAddress(1, 2, 3, 4);
            Assert.AreEqual(1, ref1.MinRow);
            Assert.AreEqual(2, ref1.MaxRow);

            ref1.FirstRow = (10);
            //now ref is CellRangeAddress(10, 2, 3, 4)
            Assert.AreEqual(2, ref1.MinRow);
            Assert.AreEqual(10, ref1.MaxRow);
        }
        [Test]
        public void TestGetMinMaxColumn()
        {
            CellRangeAddress ref1 = new CellRangeAddress(1, 2, 3, 4);
            Assert.AreEqual(3, ref1.MinColumn);
            Assert.AreEqual(4, ref1.MaxColumn);

            ref1.FirstColumn = (10);
            //now ref is CellRangeAddress(1, 2, 10, 4)
            Assert.AreEqual(4, ref1.MinColumn);
            Assert.AreEqual(10, ref1.MaxColumn);
        }

        [Test]
        public void TestIntersects()
        {
            CellRangeAddress baseRegion = new CellRangeAddress(0, 1, 0, 1);

            CellRangeAddress duplicateRegion = new CellRangeAddress(0, 1, 0, 1);
            assertIntersects(baseRegion, duplicateRegion);

            CellRangeAddress partiallyOverlappingRegion = new CellRangeAddress(1, 2, 1, 2);
            assertIntersects(baseRegion, partiallyOverlappingRegion);

            CellRangeAddress subsetRegion = new CellRangeAddress(0, 1, 0, 0);
            assertIntersects(baseRegion, subsetRegion);

            CellRangeAddress supersetRegion = new CellRangeAddress(0, 2, 0, 2);
            assertIntersects(baseRegion, supersetRegion);

            CellRangeAddress disjointRegion = new CellRangeAddress(10, 11, 10, 11);
            assertNotIntersects(baseRegion, disjointRegion);
        }


        [Test]
        public void ContainsRow()
        {
            CellRangeAddress region = new CellRangeAddress(10, 12, 3, 5);

            Assert.IsFalse(region.ContainsRow(9));
            Assert.IsTrue(region.ContainsRow(10));
            Assert.IsTrue(region.ContainsRow(11));
            Assert.IsTrue(region.ContainsRow(12));
            Assert.IsFalse(region.ContainsRow(13));
        }

        [Test]
        public void ContainsColumn()
        {
            CellRangeAddress region = new CellRangeAddress(10, 12, 3, 5);

            Assert.IsFalse(region.ContainsColumn(2));
            Assert.IsTrue(region.ContainsColumn(3));
            Assert.IsTrue(region.ContainsColumn(4));
            Assert.IsTrue(region.ContainsColumn(5));
            Assert.IsFalse(region.ContainsColumn(6));
        }



        private static void assertIntersects(CellRangeAddress regionA, CellRangeAddress regionB)
        {
            if (!(regionA.Intersects(regionB) && regionB.Intersects(regionA)))
            {
                String A = regionA.FormatAsString();
                String B = regionB.FormatAsString();
                Assert.Fail("expected: regions " + A + " and " + B + " intersect");
            }
        }
        private static void assertNotIntersects(CellRangeAddress regionA, CellRangeAddress regionB)
        {
            if ((regionA.Intersects(regionB) || regionB.Intersects(regionA)))
            {
                String A = regionA.FormatAsString();
                String B = regionB.FormatAsString();
                Assert.Fail("expected: regions " + A + " and " + B + " do not intersect");
            }
        }
    }
}