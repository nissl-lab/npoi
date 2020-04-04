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
    using System.Collections;
    using System.IO;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using TestCases.HSSF;
    using TestCases.HSSF.Record;
    using NPOI.HSSF.Record;

    /**
     * @author Marc Johnson (mjohnson at apache dot org)
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestSSTRecord
    {

        /**
         * decodes hexdump files and concatenates the results
         * @param hexDumpFileNames names of sample files in the hssf Test data directory
         */
        private static byte[] concatHexDumps(params String[] hexDumpFileNames)
        {
            int nFiles = hexDumpFileNames.Length;
            MemoryStream baos = new MemoryStream(nFiles * 8228);
            for (int i = 0; i < nFiles; i++)
            {
                String sampleFileName = hexDumpFileNames[i];
                Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(sampleFileName);
                StreamReader br = new StreamReader(is1);
                try
                {
                    while (true)
                    {
                        String line = br.ReadLine();
                        if (line == null)
                        {
                            break;
                        }
                        byte[] buffer = HexRead.ReadFromString(line);
                        baos.Write(buffer, 0, buffer.Length);
                    }
                    is1.Close();
                }
                catch (IOException e)
                {
                    throw new RuntimeException(e);
                }
            }

            return baos.ToArray();
        }

        /**
         * @param rawData serialization of one {@link SSTRecord} and zero or more {@link ContinueRecord}s
         */
        private static SSTRecord CreateSSTFromRawData(byte[] rawData)
        {
            RecordInputStream in1 = new RecordInputStream(new MemoryStream(rawData));
            in1.NextRecord();
            SSTRecord result = new SSTRecord(in1);
            Assert.AreEqual(0, in1.Remaining);
            Assert.IsTrue(!in1.HasNextRecord);
            return result;
        }

        /**
         * SST is often split over several {@link ContinueRecord}s
         */
        [Test]
        public void TestContinuedRecord()
        {
            byte[] origData;
            SSTRecord record;
            byte[] ser_output;

            origData = concatHexDumps("BigSSTRecord", "BigSSTRecordCR");
            record = CreateSSTFromRawData(origData);
            Assert.AreEqual(1464, record.NumStrings);
            Assert.AreEqual(688, record.NumUniqueStrings);
            Assert.AreEqual(688, record.CountStrings);
            ser_output = record.Serialize();
            Assert.IsTrue(Arrays.Equals(origData, ser_output));

            // Testing based on new bug report
            origData = concatHexDumps("BigSSTRecord2", "BigSSTRecord2CR1", "BigSSTRecord2CR2", "BigSSTRecord2CR3",
                    "BigSSTRecord2CR4", "BigSSTRecord2CR5", "BigSSTRecord2CR6", "BigSSTRecord2CR7");
            record = CreateSSTFromRawData(origData);


            Assert.AreEqual(158642, record.NumStrings);
            Assert.AreEqual(5249, record.NumUniqueStrings);
            Assert.AreEqual(5249, record.CountStrings);
            ser_output = record.Serialize();
#if !HIDE_UNREACHABLE_CODE
            if (false)
            { // Set true to observe make sure areSameSSTs() is working
                ser_output[11000] = (byte)'X';
            }
#endif

            SSTRecord rec2 = CreateSSTFromRawData(ser_output);
            if (!areSameSSTs(record, rec2))
            {
                throw new AssertionException("large SST re-Serialized incorrectly");
            }
#if !HIDE_UNREACHABLE_CODE
            if (false)
            {
                // TODO - trivial differences in ContinueRecord break locations
                // Sample data should be Checked against what most recent Excel version produces.
                // maybe tweaks are required in ContinuableRecordOutput
                Assert.IsTrue(Arrays.Equals(origData, ser_output));
            }
#endif
        }

        private bool areSameSSTs(SSTRecord a, SSTRecord b)
        {

            if (a.NumStrings != b.NumStrings)
            {
                return false;
            }
            int nElems = a.NumUniqueStrings;
            if (nElems != b.NumUniqueStrings)
            {
                return false;
            }
            for (int i = 0; i < nElems; i++)
            {
                if (!a.GetString(i).Equals(b.GetString(i)))
                {
                    return false;
                }
            }
            return true;
        }
        private char[] ConvertByteToChar(byte[] b)
        {
            char[] c = new char[b.Length];
            for (int i = 0; i < c.Length; i++)
            {
                c[i] = (char)b[i];
            }
            return c;
        }
        /**
         * Test capability of handling mondo big strings
         *
         * @exception IOException
         */
        [Test]
        public void TestHugeStrings()
        {
            SSTRecord record = new SSTRecord();
            byte[][] bstrings =
                {
                    new byte[9000], new byte[7433], new byte[9002],
                    new byte[16998]
                };
            UnicodeString[] strings = new UnicodeString[bstrings.Length];
            int total_length = 0;

            for (int k = 0; k < bstrings.Length; k++)
            {
                Arrays.Fill(bstrings[k], (byte)('a' + k));
                strings[k] = new UnicodeString(new String(ConvertByteToChar(bstrings[k])));
                record.AddString(strings[k]);
                total_length += 3 + bstrings[k].Length;
            }

            // add overhead of SST record
            total_length += 8;

            // add overhead of broken strings
            total_length += 4;

            // add overhead of six records
            total_length += (6 * 4);
            byte[] content = new byte[record.RecordSize];

            record.Serialize(0, content);
            Assert.AreEqual(total_length, content.Length);

            //Deserialize the record.
            RecordInputStream recStream = new RecordInputStream(new MemoryStream(content));
            recStream.NextRecord();
            record = new SSTRecord(recStream);

            Assert.AreEqual(strings.Length, record.NumStrings);
            Assert.AreEqual(strings.Length, record.NumUniqueStrings);
            Assert.AreEqual(strings.Length, record.CountStrings);
            for (int k = 0; k < strings.Length; k++)
            {
                Assert.AreEqual(strings[k], record.GetString(k));
            }
            record = new SSTRecord();
            bstrings[1] = new byte[bstrings[1].Length - 1];
            for (int k = 0; k < bstrings.Length; k++)
            {
                if ((bstrings[k].Length % 2) == 1)
                {
                    Arrays.Fill(bstrings[k], (byte)('a' + k));
                    strings[k] = new UnicodeString(new String(ConvertByteToChar(bstrings[k])));
                }
                else
                {
                    char[] data = new char[bstrings[k].Length / 2];

                    Arrays.Fill(data, (char)('\u2122' + k));
                    strings[k] = new UnicodeString(new String(data));
                }
                record.AddString(strings[k]);
            }
            content = new byte[record.RecordSize];
            record.Serialize(0, content);
            total_length--;
            Assert.AreEqual(total_length, content.Length);

            recStream = new RecordInputStream(new MemoryStream(content));
            recStream.NextRecord();
            record = new SSTRecord(recStream);

            Assert.AreEqual(strings.Length, record.NumStrings);
            Assert.AreEqual(strings.Length, record.NumUniqueStrings);
            Assert.AreEqual(strings.Length, record.CountStrings);
            for (int k = 0; k < strings.Length; k++)
            {
                Assert.AreEqual(strings[k], record.GetString(k));
            }
        }

        /**
         * Test SSTRecord boundary conditions
         */
        [Test]
        public void TestSSTRecordBug()
        {
            // create an SSTRecord and write a certain pattern of strings
            // to it ... then serialize it and verify the content
            SSTRecord record = new SSTRecord();

            // the record will start with two integers, then this string
            // ... that will eat up 16 of the 8224 bytes that the record
            // can hold
            record.AddString(new UnicodeString("Hello"));

            // now we have an Additional 8208 bytes, which is an exact
            // multiple of 16 bytes
            long Testvalue = 1000000000000L;

            for (int k = 0; k < 2000; k++)
            {
                record.AddString(new UnicodeString((Testvalue + k).ToString()));
            }
            byte[] content = new byte[record.RecordSize];

            record.Serialize(0, content);
            Assert.AreEqual(8224, LittleEndian.GetShort(content, 2));
            Assert.AreEqual(ContinueRecord.sid, LittleEndian.GetShort(content, 8228));
            Assert.AreEqual(8224, LittleEndian.GetShort(content, 8228 + 2));
            Assert.AreEqual((byte)13, content[4 + 8228]);
            Assert.AreEqual(ContinueRecord.sid, LittleEndian.GetShort(content, 2 * 8228));
            Assert.AreEqual(8224, LittleEndian.GetShort(content, 8228 * 2 + 2));
            Assert.AreEqual((byte)13, content[4 + 8228 * 2]);
            Assert.AreEqual(ContinueRecord.sid, LittleEndian.GetShort(content, 3 * 8228));
            Assert.AreEqual((byte)13, content[4 + 8228 * 3]);
        }

        /**
         * Test simple AddString
         */
        [Test]
        public void TestSimpleAddString()
        {
            SSTRecord record = new SSTRecord();
            UnicodeString s1 = new UnicodeString("Hello world");

            // \u2122 is the encoding of the trademark symbol ...
            UnicodeString s2 = new UnicodeString("Hello world\u2122");

            Assert.AreEqual(0, record.AddString(s1));
            Assert.AreEqual(s1, record.GetString(0));
            Assert.AreEqual(1, record.CountStrings);
            Assert.AreEqual(1, record.NumStrings);
            Assert.AreEqual(1, record.NumUniqueStrings);
            Assert.AreEqual(0, record.AddString(s1));
            Assert.AreEqual(s1, record.GetString(0));
            Assert.AreEqual(1, record.CountStrings);
            Assert.AreEqual(2, record.NumStrings);
            Assert.AreEqual(1, record.NumUniqueStrings);
            Assert.AreEqual(1, record.AddString(s2));
            Assert.AreEqual(s2, record.GetString(1));
            Assert.AreEqual(2, record.CountStrings);
            Assert.AreEqual(3, record.NumStrings);
            Assert.AreEqual(2, record.NumUniqueStrings);
            IEnumerator iter = record.GetStrings();

            while (iter.MoveNext())
            {
                UnicodeString ucs = (UnicodeString)iter.Current;

                if (ucs.Equals(s1))
                {
                    Assert.AreEqual((byte)0, ucs.OptionFlags);
                }
                else if (ucs.Equals(s2))
                {
                    Assert.AreEqual((byte)1, ucs.OptionFlags);
                }
                else
                {
                    Assert.Fail("cannot match string: " + ucs.String);
                }
            }
        }

        /**
         * Test simple constructor
         */
        [Test]
        public void TestSimpleConstructor()
        {
            SSTRecord record = new SSTRecord();

            Assert.AreEqual(0, record.NumStrings);
            Assert.AreEqual(0, record.NumUniqueStrings);
            Assert.AreEqual(0, record.CountStrings);
            byte[] output = record.Serialize();
            byte[] expected =
                {
                    (byte) record.Sid, (byte) ( record.Sid >> 8 ),
                    (byte) 8, (byte) 0, (byte) 0, (byte) 0, (byte) 0,
                    (byte) 0, (byte) 0, (byte) 0, (byte) 0, (byte) 0
                };

            Assert.AreEqual(expected.Length, output.Length);
            for (int k = 0; k < expected.Length; k++)
            {
                Assert.AreEqual(expected[k], output[k], k.ToString());
            }
        }

        /**
         * Tests that workbooks with rich text that duplicates a non rich text cell can be read and written.
         */
        [Test]
        public void TestReadWriteDuplicatedRichText1()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("duprich1.xls");
            ISheet sheet = wb.GetSheetAt(1);
            Assert.AreEqual("01/05 (Wed)", sheet.GetRow(0).GetCell(8).StringCellValue);
            Assert.AreEqual("01/05 (Wed)", sheet.GetRow(1).GetCell(8).StringCellValue);

            HSSFTestDataSamples.WriteOutAndReadBack((HSSFWorkbook)wb).Close();
            wb.Close();
            // Test the second file.
            wb = HSSFTestDataSamples.OpenSampleWorkbook("duprich2.xls");
            sheet = wb.GetSheetAt(0);
            int row = 0;
            Assert.AreEqual("Testing", sheet.GetRow(row++).GetCell(0).StringCellValue);
            Assert.AreEqual("rich", sheet.GetRow(row++).GetCell(0).StringCellValue);
            Assert.AreEqual("text", sheet.GetRow(row++).GetCell(0).StringCellValue);
            Assert.AreEqual("strings", sheet.GetRow(row++).GetCell(0).StringCellValue);
            Assert.AreEqual("Testing", sheet.GetRow(row++).GetCell(0).StringCellValue);
            Assert.AreEqual("Testing", sheet.GetRow(row++).GetCell(0).StringCellValue);

            HSSFTestDataSamples.WriteOutAndReadBack((HSSFWorkbook)wb).Close();
            wb.Close();
        }

        /**
         * hex dump from UnicodeStringFailCase1.xls atatched to Bugzilla 50779
         */
        private static String data_50779_1 =
            //Offset=0x00000612(1554) recno=71 sid=0x00FC size=0x2020(8224)
                "      FC 00 20 20 51 00 00 00 51 00 00 00 32 00" +
                "05 10 00 00 00 42 30 44 30 46 30 48 30 4A 30 4B" +
                "30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B" +
                "30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B" +
                "30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B" +
                "30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88" +
                "30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93" +
                "30 30 00 30 00 30 00 31 00 01 00 0C 00 05 00 35" +
                "00 00 00 00 00 00 00 4B 30 32 00 01 42 30 44 30" +
                "46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30" +
                "55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30" +
                "66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30" +
                "72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30" +
                "82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30" +
                "8D 30 8F 30 92 30 93 30 30 00 30 00 30 00 32 00" +
                "32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D" +
                "30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D" +
                "30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C" +
                "30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E" +
                "30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89" +
                "30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30" +
                "00 30 00 30 00 33 00 32 00 01 42 30 44 30 46 30" +
                "48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30" +
                "57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30" +
                "68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30" +
                "75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30" +
                "84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30" +
                "8F 30 92 30 93 30 30 00 30 00 30 00 34 00 32 00" +
                "01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F" +
                "30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F" +
                "30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D" +
                "30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F" +
                "30 80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A" +
                "30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30" +
                "00 30 00 35 00 32 00 01 42 30 44 30 46 30 48 30" +
                "4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30" +
                "59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30" +
                "6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30" +
                "78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30" +
                "86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30" +
                "92 30 93 30 30 00 30 00 30 00 36 00 32 00 01 42" +
                "30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51" +
                "30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61" +
                "30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E" +
                "30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80" +
                "30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B" +
                "30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 30" +
                "00 37 00 32 00 01 42 30 44 30 46 30 48 30 4A 30" +
                "4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30" +
                "5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30" +
                "6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30" +
                "7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30" +
                "88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30" +
                "93 30 30 00 30 00 30 00 38 00 32 00 01 42 30 44" +
                "30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53" +
                "30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64" +
                "30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F" +
                "30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81" +
                "30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C" +
                "30 8D 30 8F 30 92 30 93 30 30 00 30 00 30 00 39" +
                "00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30" +
                "4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30" +
                "5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30" +
                "6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30" +
                "7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30" +
                "89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30" +
                "30 00 30 00 31 00 30 00 32 00 01 42 30 44 30 46" +
                "30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55" +
                "30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66" +
                "30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72" +
                "30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82" +
                "30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D" +
                "30 8F 30 92 30 93 30 30 00 30 00 31 00 31 00 32" +
                "00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30" +
                "4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30" +
                "5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30" +
                "6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30" +
                "7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89 30" +
                "8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00" +
                "30 00 31 00 32 00 32 00 01 42 30 44 30 46 30 48" +
                "30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57" +
                "30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68" +
                "30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75" +
                "30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84" +
                "30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F" +
                "30 92 30 93 30 30 00 30 00 31 00 33 00 32 00 01" +
                "42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30" +
                "51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30" +
                "61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30" +
                "6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30" +
                "80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30" +
                "8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00" +
                "31 00 34 00 32 00 01 42 30 44 30 46 30 48 30 4A" +
                "30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59" +
                "30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A" +
                "30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78" +
                "30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86" +
                "30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92" +
                "30 93 30 30 00 30 00 31 00 35 00 32 00 01 42 30" +
                "44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30" +
                "53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30" +
                "64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30" +
                "6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30" +
                "81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30" +
                "8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 31 00" +
                "36 00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B" +
                "30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B" +
                "30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B" +
                "30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B" +
                "30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88" +
                "30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93" +
                "30 30 00 30 00 31 00 37 00 32 00 01 42 30 44 30" +
                "46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30" +
                "55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30" +
                "66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30" +
                "72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30" +
                "82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30" +
                "8D 30 8F 30 92 30 93 30 30 00 30 00 31 00 38 00" +
                "32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D" +
                "30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D" +
                "30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C" +
                "30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E" +
                "30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89" +
                "30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30" +
                "00 30 00 31 00 39 00 32 00 01 42 30 44 30 46 30" +
                "48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30" +
                "57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30" +
                "68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30" +
                "75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30" +
                "84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30" +
                "8F 30 92 30 93 30 30 00 30 00 32 00 30 00 32 00" +
                "01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F" +
                "30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F" +
                "30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D" +
                "30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F" +
                "30 80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A" +
                "30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30" +
                "00 32 00 31 00 32 00 01 42 30 44 30 46 30 48 30" +
                "4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30" +
                "59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30" +
                "6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30" +
                "78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30" +
                "86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30" +
                "92 30 93 30 30 00 30 00 32 00 32 00 32 00 01 42" +
                "30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51" +
                "30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61" +
                "30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E" +
                "30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80" +
                "30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B" +
                "30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 32" +
                "00 33 00 32 00 01 42 30 44 30 46 30 48 30 4A 30" +
                "4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30" +
                "5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30" +
                "6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30" +
                "7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30" +
                "88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30" +
                "93 30 30 00 30 00 32 00 34 00 32 00 01 42 30 44" +
                "30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53" +
                "30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64" +
                "30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F" +
                "30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81" +
                "30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C" +
                "30 8D 30 8F 30 92 30 93 30 30 00 30 00 32 00 35" +
                "00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30" +
                "4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30" +
                "5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30" +
                "6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30" +
                "7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30" +
                "89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30" +
                "30 00 30 00 32 00 36 00 32 00 01 42 30 44 30 46" +
                "30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55" +
                "30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66" +
                "30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72" +
                "30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82" +
                "30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D" +
                "30 8F 30 92 30 93 30 30 00 30 00 32 00 37 00 32" +
                "00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30" +
                "4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30" +
                "5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30" +
                "6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30" +
                "7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89 30" +
                "8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00" +
                "30 00 32 00 38 00 32 00 01 42 30 44 30 46 30 48" +
                "30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57" +
                "30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68" +
                "30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75" +
                "30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84" +
                "30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F" +
                "30 92 30 93 30 30 00 30 00 32 00 39 00 32 00 01" +
                "42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30" +
                "51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30" +
                "61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30" +
                "6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30" +
                "80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30" +
                "8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00" +
                "33 00 30 00 32 00 01 42 30 44 30 46 30 48 30 4A" +
                "30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59" +
                "30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A" +
                "30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78" +
                "30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86" +
                "30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92" +
                "30 93 30 30 00 30 00 33 00 31 00 32 00 01 42 30" +
                "44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30" +
                "53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30" +
                "64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30" +
                "6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30" +
                "81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30" +
                "8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 33 00" +
                "32 00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B" +
                "30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B" +
                "30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B" +
                "30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B" +
                "30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88" +
                "30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93" +
                "30 30 00 30 00 33 00 33 00 32 00 01 42 30 44 30" +
                "46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30" +
                "55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30" +
                "66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30" +
                "72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30" +
                "82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30" +
                "8D 30 8F 30 92 30 93 30 30 00 30 00 33 00 34 00" +
                "32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D" +
                "30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D" +
                "30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C" +
                "30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E" +
                "30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89" +
                "30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30" +
                "00 30 00 33 00 35 00 32 00 01 42 30 44 30 46 30" +
                "48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30" +
                "57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30" +
                "68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30" +
                "75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30" +
                "84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30" +
                "8F 30 92 30 93 30 30 00 30 00 33 00 36 00 32 00" +
                "01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F" +
                "30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F" +
                "30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D" +
                "30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F" +
                "30 80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A" +
                "30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30" +
                "00 33 00 37 00 32 00 01 42 30 44 30 46 30 48 30" +
                "4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30" +
                "59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30" +
                "6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30" +
                "78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30" +
                "86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30" +
                "92 30 93 30 30 00 30 00 33 00 38 00 32 00 01 42" +
                "30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51" +
                "30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61" +
                "30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E" +
                "30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80" +
                "30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B" +
                "30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 33" +
                "00 39 00 32 00 01 42 30 44 30 46 30 48 30 4A 30" +
                "4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30" +
                "5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30" +
                "6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30" +
                "7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30" +
                "88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30" +
                "93 30 30 00 30 00 34 00 30 00 32 00 01 42 30 44" +
                "30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53" +
                "30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64" +
                "30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F" +
                "30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81" +
                "30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C" +
                "30 8D 30 8F 30 92 30 93 30 30 00 30 00 34 00 31" +
                "00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30" +
                "4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30" +
                "5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30" +
                "6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30" +
                "7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30" +
                "89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30" +
                "30 00 30 00 34 00 32 00 32 00 01 42 30 44 30 46" +
                "30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55" +
                "30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66" +
                "30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72" +
                "30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82" +
                "30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D" +
                "30 8F 30 92 30 93 30 30 00 30 00 34 00 33 00 32" +
                "00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30" +
                "4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30" +
                "5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30" +
                "6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30" +
                "7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89 30" +
                "8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00" +
                "30 00 34 00 34 00 32 00 01 42 30 44 30 46 30 48" +
                "30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57" +
                "30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68" +
                "30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75" +
                "30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84" +
                "30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F" +
                "30 92 30 93 30 30 00 30 00 34 00 35 00 32 00 01" +
                "42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30" +
                "51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30" +
                "61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30" +
                "6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30" +
                "80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30" +
                "8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00" +
                "34 00 36 00 32 00 01 42 30 44 30 46 30 48 30 4A" +
                "30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59" +
                "30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A" +
                "30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78" +
                "30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86" +
                "30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92" +
                "30 93 30 30 00 30 00 34 00 37 00 32 00 01 42 30" +
                "44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30" +
                "53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30" +
                "64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30" +
                "6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30" +
                "81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30" +
                "8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 34 00" +
                "38 00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B" +
                "30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B" +
                "30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B" +
                "30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B" +
                "30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88" +
                "30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93" +
                "30 30 00 30 00 34 00 39 00 32 00 01 42 30 44 30" +
                "46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30" +
                "55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30" +
                "66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30" +
                "72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30" +
                "82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30" +
                "8D 30 8F 30 92 30 93 30 30 00 30 00 35 00 30 00" +
                "32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D" +
                "30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D" +
                "30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C" +
                "30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E" +
                "30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89" +
                "30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30" +
                "00 30 00 35 00 31 00 32 00 01 42 30 44 30 46 30" +
                "48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30" +
                "57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30" +
                "68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30" +
                "75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30" +
                "84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30" +
                "8F 30 92 30 93 30 30 00 30 00 35 00 32 00 32 00" +
                "01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F" +
                "30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F" +
                "30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D" +
                "30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F" +
                "30 80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A" +
                "30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30" +
                "00 35 00 33 00 32 00 01 42 30 44 30 46 30 48 30" +
                "4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30" +
                "59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30" +
                "6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30" +
                "78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30" +
                "86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30" +
                "92 30 93 30 30 00 30 00 35 00 34 00 32 00 01 42" +
                "30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51" +
                "30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61" +
                "30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E" +
                "30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80" +
                "30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B" +
                "30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 35" +
                "00 35 00 32 00 01 42 30 44 30 46 30 48 30 4A 30" +
                "4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30" +
                "5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30" +
                "6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30" +
                "7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30" +
                "88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30" +
                "93 30 30 00 30 00 35 00 36 00 32 00 01 42 30 44" +
                "30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53" +
                "30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64" +
                "30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F" +
                "30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81" +
                "30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C" +
                "30 8D 30 8F 30 92 30 93 30 30 00 30 00 35 00 37" +
                "00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30" +
                "4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30" +
                "5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30" +
                "6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30" +
                "7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30" +
                "89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30" +
                "30 00 30 00 35 00 38 00 32 00 01 42 30 44 30 46" +
                "30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55" +
                "30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66" +
                "30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72" +
                "30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82" +
                "30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D" +
                "30 8F 30 92 30 93 30 30 00 30 00 35 00 39 00 32" +
                "00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30" +
                "4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30" +
                "5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30" +
                "6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30" +
                "7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89 30" +
                "8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00" +
                "30 00 36 00 30 00 32 00 01 42 30 44 30 46 30 48" +
                "30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57" +
                "30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68" +
                "30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75" +
                "30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84" +
                "30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F" +
                "30 92 30 93 30 30 00 30 00 36 00 31 00 32 00 01" +
                "42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30" +
                "51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30" +
                "61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30" +
                "6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30" +
                "80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30" +
                "8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00" +
                "36 00 32 00 32 00 01 42 30 44 30 46 30 48 30 4A" +
                "30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59" +
                "30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A" +
                "30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78" +
                "30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86" +
                "30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92" +
                "30 93 30 30 00 30 00 36 00 33 00 32 00 01 42 30" +
                "44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30" +
                "53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30" +
                "64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30" +
                "6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30" +
                "81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30" +
                "8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 36 00" +
                "34 00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B" +
                "30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B" +
                "30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B" +
                "30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B" +
                "30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88" +
                "30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93" +
                "30 30 00 30 00 36 00 35 00 32 00 01 42 30 44 30" +
                "46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30" +
                "55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30" +
                "66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30" +
                "72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30" +
                "82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30" +
                "8D 30 8F 30 92 30 93 30 30 00 30 00 36 00 36 00" +
                "32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D" +
                "30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D" +
                "30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C" +
                "30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E" +
                "30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89" +
                "30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30" +
                "00 30 00 36 00 37 00 32 00 01 42 30 44 30 46 30" +
                "48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30" +
                "57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30" +
                "68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30" +
                "75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30" +
                "84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30" +
                "8F 30 92 30 93 30 30 00 30 00 36 00 38 00 32 00" +
                "01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F" +
                "30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F" +
                "30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D" +
                "30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F" +
                "30 80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A" +
                "30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30" +
                "00 36 00 39 00 32 00 01 42 30 44 30 46 30 48 30" +
                "4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30" +
                "59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30" +
                "6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30" +
                "78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30" +
                "86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30" +
                "92 30 93 30 30 00 30 00 37 00 30 00 32 00 01 42" +
                "30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51" +
                "30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61" +
                "30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E" +
                "30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80" +
                "30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B" +
                "30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 37" +
                "00 31 00 32 00 01 42 30 44 30 46 30 48 30 4A 30" +
                "4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30" +
                "5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30" +
                "6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30" +
                "7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30" +
                "88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30" +
                "93 30 30 00 30 00 37 00 32 00 32 00 01 42 30 44" +
                "30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53" +
                "30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64" +
                "30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F" +
                "30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81" +
                "30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C" +
                "30 8D 30 8F 30 92 30 93 30 30 00 30 00 37 00 33" +
                "00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30" +
                "4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30" +
                "5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30" +
                "6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30" +
                "7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30" +
                "89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30" +
                "30 00 30 00 37 00 34 00 32 00 01 42 30 44 30 46" +
                "30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55" +
                "30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66" +
                "30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72" +
                "30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82" +
                "30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D" +
                "30 8F 30 92 30 93 30 30 00 30 00 37 00 35 00 32" +
                "00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30" +
                "4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30" +
                "5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30" +
                "6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30" +
                "7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89 30" +
                "8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00" +
                "30 00 37 00 36 00 32 00 01 42 30 44 30 46 30 48" +
                "30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57" +
                "30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68" +
                "30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75" +
                "30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84" +
                "30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F" +
                "30 92 30 93 30 30 00 30 00 37 00 37 00 32 00 01" +
                "42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30" +
                "51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30" +
                "61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30" +
                "6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30" +
                "80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30" +
                "8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00" +
                "37 00 38 00 1F 00 05 B8 00 00 00 17 53 77 6D 53" +
                "90 52 97 EE 68 0C 77 A9 5C 4B 62 0C 77 8F 79 F6" +
                "5C 0C 77 03 68 28 67 0C 77 FC 57 89 73 0C 77 71" +
                "67 AC 4E FD 90 43 53 49 84 0C 77 5E 79 48 59 DD" +
                "5D 0C 77 77 95 CE 91 0C 77 01 00 B4 00 05 00 35" +
                "00 0A 00 37 00 37 00 DB 30 C3 30 AB 30 A4 30 C9" +
                "30 A6 30 A2 30 AA 30 E2 30 EA 30 B1 30 F3 30 A4" +
                "30 EF 30 C6 30 B1 30 F3 30 D5 30 AF 30 B7 30 DE" +
                "30 B1 30 F3 30 C8 30 C1 30 AE 30 B1 30 F3 30 B5" +
                "30 A4 30 BF 30 DE 30 B1 30 F3 30 C8 30 A6 30 AD" +
                "30 E7 30 A6 30 C8                              " +

                // Offset=0x00002636(9782) recno=72 sid=0x003C size=0x0151(337)
                "                  3C 00 51 01 30 C1 30 D0 30 B1" +
                "30 F3 30 AB 30 CA 30 AC 30 EF 30 B1 30 F3 30 CA" +
                "30 AC 30 CE 30 B1 30 F3 30 00 00 00 00 03 00 06" +
                "00 03 00 03 00 0C 00 06 00 03 00 11 00 09 00 03" +
                "00 17 00 0C 00 03 00 1C 00 0F 00 03 00 22 00 12" +
                "00 03 00 28 00 15 00 03 00 2C 00 18 00 04 00 32" +
                "00 1C 00 03 00 32 00 05 10 00 00 00 42 30 44 30" +
                "46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30" +
                "55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30" +
                "66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30" +
                "72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30" +
                "82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30" +
                "8D 30 8F 30 92 30 93 30 30 00 30 00 37 00 39 00" +
                "01 00 0C 00 05 00 35 00 00 00 00 00 00 00 00 00" +
                "32 00 05 10 00 00 00 42 30 44 30 46 30 48 30 4A" +
                "30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59" +
                "30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A" +
                "30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78" +
                "30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86" +
                "30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92" +
                "30 93 30 30 00 30 00 38 00 30 00 01 00 0C 00 05" +
                "00 35 00 00 00 00 00 00 00 4B 30               ";


        /**
         * hex dump from UnicodeStringFailCase2.xls atatched to Bugzilla 50779
         */
        private static String data_50779_2 =
            //"Offset=0x00000612(1554) recno=71 sid=0x00FC size=0x2020(8224)\n" +
                "      FC 00 20 20 51 00 00 00 51 00 00 00 32 00" +
                "01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F" +
                "30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F" +
                "30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D" +
                "30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F" +
                "30 80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A" +
                "30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30" +
                "00 30 00 31 00 32 00 01 42 30 44 30 46 30 48 30" +
                "4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30" +
                "59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30" +
                "6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30" +
                "78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30" +
                "86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30" +
                "92 30 93 30 30 00 30 00 30 00 32 00 32 00 01 42" +
                "30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51" +
                "30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61" +
                "30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E" +
                "30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80" +
                "30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B" +
                "30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 30" +
                "00 33 00 32 00 01 42 30 44 30 46 30 48 30 4A 30" +
                "4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30" +
                "5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30" +
                "6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30" +
                "7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30" +
                "88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30" +
                "93 30 30 00 30 00 30 00 34 00 32 00 01 42 30 44" +
                "30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53" +
                "30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64" +
                "30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F" +
                "30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81" +
                "30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C" +
                "30 8D 30 8F 30 92 30 93 30 30 00 30 00 30 00 35" +
                "00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30" +
                "4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30" +
                "5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30" +
                "6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30" +
                "7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30" +
                "89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30" +
                "30 00 30 00 30 00 36 00 32 00 01 42 30 44 30 46" +
                "30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55" +
                "30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66" +
                "30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72" +
                "30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82" +
                "30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D" +
                "30 8F 30 92 30 93 30 30 00 30 00 30 00 37 00 32" +
                "00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30" +
                "4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30" +
                "5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30" +
                "6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30" +
                "7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89 30" +
                "8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00" +
                "30 00 30 00 38 00 32 00 01 42 30 44 30 46 30 48" +
                "30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57" +
                "30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68" +
                "30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75" +
                "30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84" +
                "30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F" +
                "30 92 30 93 30 30 00 30 00 30 00 39 00 32 00 01" +
                "42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30" +
                "51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30" +
                "61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30" +
                "6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30" +
                "80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30" +
                "8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00" +
                "31 00 30 00 32 00 01 42 30 44 30 46 30 48 30 4A" +
                "30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59" +
                "30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A" +
                "30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78" +
                "30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86" +
                "30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92" +
                "30 93 30 30 00 30 00 31 00 31 00 32 00 01 42 30" +
                "44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30" +
                "53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30" +
                "64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30" +
                "6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30" +
                "81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30" +
                "8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 31 00" +
                "32 00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B" +
                "30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B" +
                "30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B" +
                "30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B" +
                "30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88" +
                "30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93" +
                "30 30 00 30 00 31 00 33 00 32 00 01 42 30 44 30" +
                "46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30" +
                "55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30" +
                "66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30" +
                "72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30" +
                "82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30" +
                "8D 30 8F 30 92 30 93 30 30 00 30 00 31 00 34 00" +
                "32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D" +
                "30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D" +
                "30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C" +
                "30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E" +
                "30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89" +
                "30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30" +
                "00 30 00 31 00 35 00 32 00 01 42 30 44 30 46 30" +
                "48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30" +
                "57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30" +
                "68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30" +
                "75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30" +
                "84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30" +
                "8F 30 92 30 93 30 30 00 30 00 31 00 36 00 32 00" +
                "01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F" +
                "30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F" +
                "30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D" +
                "30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F" +
                "30 80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A" +
                "30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30" +
                "00 31 00 37 00 32 00 01 42 30 44 30 46 30 48 30" +
                "4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30" +
                "59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30" +
                "6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30" +
                "78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30" +
                "86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30" +
                "92 30 93 30 30 00 30 00 31 00 38 00 32 00 01 42" +
                "30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51" +
                "30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61" +
                "30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E" +
                "30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80" +
                "30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B" +
                "30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 31" +
                "00 39 00 32 00 01 42 30 44 30 46 30 48 30 4A 30" +
                "4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30" +
                "5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30" +
                "6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30" +
                "7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30" +
                "88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30" +
                "93 30 30 00 30 00 32 00 30 00 32 00 01 42 30 44" +
                "30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53" +
                "30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64" +
                "30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F" +
                "30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81" +
                "30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C" +
                "30 8D 30 8F 30 92 30 93 30 30 00 30 00 32 00 31" +
                "00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30" +
                "4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30" +
                "5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30" +
                "6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30" +
                "7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30" +
                "89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30" +
                "30 00 30 00 32 00 32 00 32 00 01 42 30 44 30 46" +
                "30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55" +
                "30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66" +
                "30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72" +
                "30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82" +
                "30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D" +
                "30 8F 30 92 30 93 30 30 00 30 00 32 00 33 00 32" +
                "00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30" +
                "4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30" +
                "5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30" +
                "6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30" +
                "7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89 30" +
                "8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00" +
                "30 00 32 00 34 00 32 00 01 42 30 44 30 46 30 48" +
                "30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57" +
                "30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68" +
                "30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75" +
                "30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84" +
                "30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F" +
                "30 92 30 93 30 30 00 30 00 32 00 35 00 32 00 01" +
                "42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30" +
                "51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30" +
                "61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30" +
                "6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30" +
                "80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30" +
                "8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00" +
                "32 00 36 00 32 00 01 42 30 44 30 46 30 48 30 4A" +
                "30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59" +
                "30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A" +
                "30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78" +
                "30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86" +
                "30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92" +
                "30 93 30 30 00 30 00 32 00 37 00 32 00 01 42 30" +
                "44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30" +
                "53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30" +
                "64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30" +
                "6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30" +
                "81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30" +
                "8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 32 00" +
                "38 00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B" +
                "30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B" +
                "30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B" +
                "30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B" +
                "30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88" +
                "30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93" +
                "30 30 00 30 00 32 00 39 00 32 00 01 42 30 44 30" +
                "46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30" +
                "55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30" +
                "66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30" +
                "72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30" +
                "82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30" +
                "8D 30 8F 30 92 30 93 30 30 00 30 00 33 00 30 00" +
                "32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D" +
                "30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D" +
                "30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C" +
                "30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E" +
                "30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89" +
                "30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30" +
                "00 30 00 33 00 31 00 32 00 01 42 30 44 30 46 30" +
                "48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30" +
                "57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30" +
                "68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30" +
                "75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30" +
                "84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30" +
                "8F 30 92 30 93 30 30 00 30 00 33 00 32 00 32 00" +
                "01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F" +
                "30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F" +
                "30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D" +
                "30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F" +
                "30 80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A" +
                "30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30" +
                "00 33 00 33 00 32 00 01 42 30 44 30 46 30 48 30" +
                "4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30" +
                "59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30" +
                "6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30" +
                "78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30" +
                "86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30" +
                "92 30 93 30 30 00 30 00 33 00 34 00 32 00 01 42" +
                "30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51" +
                "30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61" +
                "30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E" +
                "30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80" +
                "30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B" +
                "30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 33" +
                "00 35 00 32 00 01 42 30 44 30 46 30 48 30 4A 30" +
                "4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30" +
                "5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30" +
                "6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30" +
                "7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30" +
                "88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30" +
                "93 30 30 00 30 00 33 00 36 00 32 00 01 42 30 44" +
                "30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53" +
                "30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64" +
                "30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F" +
                "30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81" +
                "30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C" +
                "30 8D 30 8F 30 92 30 93 30 30 00 30 00 33 00 37" +
                "00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30" +
                "4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30" +
                "5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30" +
                "6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30" +
                "7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30" +
                "89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30" +
                "30 00 30 00 33 00 38 00 32 00 01 42 30 44 30 46" +
                "30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55" +
                "30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66" +
                "30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72" +
                "30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82" +
                "30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D" +
                "30 8F 30 92 30 93 30 30 00 30 00 33 00 39 00 32" +
                "00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30" +
                "4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30" +
                "5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30" +
                "6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30" +
                "7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89 30" +
                "8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00" +
                "30 00 34 00 30 00 32 00 01 42 30 44 30 46 30 48" +
                "30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57" +
                "30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68" +
                "30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75" +
                "30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84" +
                "30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F" +
                "30 92 30 93 30 30 00 30 00 34 00 31 00 32 00 01" +
                "42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30" +
                "51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30" +
                "61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30" +
                "6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30" +
                "80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30" +
                "8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00" +
                "34 00 32 00 32 00 01 42 30 44 30 46 30 48 30 4A" +
                "30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59" +
                "30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A" +
                "30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78" +
                "30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86" +
                "30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92" +
                "30 93 30 30 00 30 00 34 00 33 00 32 00 01 42 30" +
                "44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30" +
                "53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30" +
                "64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30" +
                "6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30" +
                "81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30" +
                "8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 34 00" +
                "34 00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B" +
                "30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B" +
                "30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B" +
                "30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B" +
                "30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88" +
                "30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93" +
                "30 30 00 30 00 34 00 35 00 32 00 01 42 30 44 30" +
                "46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30" +
                "55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30" +
                "66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30" +
                "72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30" +
                "82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30" +
                "8D 30 8F 30 92 30 93 30 30 00 30 00 34 00 36 00" +
                "32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D" +
                "30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D" +
                "30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C" +
                "30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E" +
                "30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89" +
                "30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30" +
                "00 30 00 34 00 37 00 32 00 01 42 30 44 30 46 30" +
                "48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30" +
                "57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30" +
                "68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30" +
                "75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30" +
                "84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30" +
                "8F 30 92 30 93 30 30 00 30 00 34 00 38 00 32 00" +
                "01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F" +
                "30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F" +
                "30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D" +
                "30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F" +
                "30 80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A" +
                "30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30" +
                "00 34 00 39 00 32 00 01 42 30 44 30 46 30 48 30" +
                "4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30" +
                "59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30" +
                "6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30" +
                "78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30" +
                "86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30" +
                "92 30 93 30 30 00 30 00 35 00 30 00 32 00 01 42" +
                "30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51" +
                "30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61" +
                "30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E" +
                "30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80" +
                "30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B" +
                "30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 35" +
                "00 31 00 32 00 01 42 30 44 30 46 30 48 30 4A 30" +
                "4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30" +
                "5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30" +
                "6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30" +
                "7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30" +
                "88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30" +
                "93 30 30 00 30 00 35 00 32 00 32 00 01 42 30 44" +
                "30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53" +
                "30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64" +
                "30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F" +
                "30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81" +
                "30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C" +
                "30 8D 30 8F 30 92 30 93 30 30 00 30 00 35 00 33" +
                "00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30" +
                "4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30" +
                "5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30" +
                "6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30" +
                "7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30" +
                "89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30" +
                "30 00 30 00 35 00 34 00 32 00 01 42 30 44 30 46" +
                "30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55" +
                "30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66" +
                "30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72" +
                "30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82" +
                "30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D" +
                "30 8F 30 92 30 93 30 30 00 30 00 35 00 35 00 32" +
                "00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30" +
                "4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30" +
                "5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30" +
                "6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30" +
                "7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89 30" +
                "8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00" +
                "30 00 35 00 36 00 32 00 01 42 30 44 30 46 30 48" +
                "30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57" +
                "30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68" +
                "30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75" +
                "30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84" +
                "30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F" +
                "30 92 30 93 30 30 00 30 00 35 00 37 00 32 00 01" +
                "42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30" +
                "51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30" +
                "61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30" +
                "6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30" +
                "80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30" +
                "8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00" +
                "35 00 38 00 32 00 01 42 30 44 30 46 30 48 30 4A" +
                "30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59" +
                "30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A" +
                "30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78" +
                "30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86" +
                "30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92" +
                "30 93 30 30 00 30 00 35 00 39 00 32 00 01 42 30" +
                "44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30" +
                "53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30" +
                "64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30" +
                "6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30" +
                "81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30" +
                "8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 36 00" +
                "30 00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B" +
                "30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B" +
                "30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B" +
                "30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B" +
                "30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88" +
                "30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93" +
                "30 30 00 30 00 36 00 31 00 32 00 01 42 30 44 30" +
                "46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30" +
                "55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30" +
                "66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30" +
                "72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30" +
                "82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30" +
                "8D 30 8F 30 92 30 93 30 30 00 30 00 36 00 32 00" +
                "32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D" +
                "30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D" +
                "30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C" +
                "30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E" +
                "30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89" +
                "30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30" +
                "00 30 00 36 00 33 00 32 00 01 42 30 44 30 46 30" +
                "48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30" +
                "57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30" +
                "68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30" +
                "75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30" +
                "84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30" +
                "8F 30 92 30 93 30 30 00 30 00 36 00 34 00 32 00" +
                "01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F" +
                "30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F" +
                "30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D" +
                "30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F" +
                "30 80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A" +
                "30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30" +
                "00 36 00 35 00 32 00 01 42 30 44 30 46 30 48 30" +
                "4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30" +
                "59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30" +
                "6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30" +
                "78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30" +
                "86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30" +
                "92 30 93 30 30 00 30 00 36 00 36 00 32 00 01 42" +
                "30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51" +
                "30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61" +
                "30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E" +
                "30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80" +
                "30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B" +
                "30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 36" +
                "00 37 00 32 00 01 42 30 44 30 46 30 48 30 4A 30" +
                "4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30" +
                "5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30" +
                "6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30" +
                "7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30" +
                "88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30" +
                "93 30 30 00 30 00 36 00 38 00 32 00 01 42 30 44" +
                "30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53" +
                "30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64" +
                "30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F" +
                "30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81" +
                "30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C" +
                "30 8D 30 8F 30 92 30 93 30 30 00 30 00 36 00 39" +
                "00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B 30" +
                "4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30" +
                "5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30" +
                "6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30" +
                "7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30" +
                "89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30" +
                "30 00 30 00 37 00 30 00 32 00 01 42 30 44 30 46" +
                "30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55" +
                "30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66" +
                "30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72" +
                "30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82" +
                "30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D" +
                "30 8F 30 92 30 93 30 30 00 30 00 37 00 31 00 32" +
                "00 01 42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30" +
                "4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D 30" +
                "5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30" +
                "6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30" +
                "7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89 30" +
                "8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00" +
                "30 00 37 00 32 00 32 00 01 42 30 44 30 46 30 48" +
                "30 4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57" +
                "30 59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68" +
                "30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75" +
                "30 78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84" +
                "30 86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F" +
                "30 92 30 93 30 30 00 30 00 37 00 33 00 32 00 01" +
                "42 30 44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30" +
                "51 30 53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30" +
                "61 30 64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30" +
                "6E 30 6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30" +
                "80 30 81 30 82 30 84 30 86 30 88 30 89 30 8A 30" +
                "8B 30 8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00" +
                "37 00 34 00 32 00 01 42 30 44 30 46 30 48 30 4A" +
                "30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30 59" +
                "30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A" +
                "30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78" +
                "30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30 86" +
                "30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92" +
                "30 93 30 30 00 30 00 37 00 35 00 32 00 01 42 30" +
                "44 30 46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30" +
                "53 30 55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30" +
                "64 30 66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30" +
                "6F 30 72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30" +
                "81 30 82 30 84 30 86 30 88 30 89 30 8A 30 8B 30" +
                "8C 30 8D 30 8F 30 92 30 93 30 30 00 30 00 37 00" +
                "36 00 32 00 01 42 30 44 30 46 30 48 30 4A 30 4B" +
                "30 4D 30 4F 30 51 30 53 30 55 30 57 30 59 30 5B" +
                "30 5D 30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B" +
                "30 6C 30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B" +
                "30 7E 30 7F 30 80 30 81 30 82 30 84 30 86 30 88" +
                "30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93" +
                "30 30 00 30 00 37 00 37 00 32 00 01 42 30 44 30" +
                "46 30 48 30 4A 30 4B 30 4D 30 4F 30 51 30 53 30" +
                "55 30 57 30 59 30 5B 30 5D 30 5F 30 61 30 64 30" +
                "66 30 68 30 6A 30 6B 30 6C 30 6D 30 6E 30 6F 30" +
                "72 30 75 30 78 30 7B 30 7E 30 7F 30 80 30 81 30" +
                "82 30 84 30 86 30 88 30 89 30 8A 30 8B 30 8C 30" +
                "8D 30 8F 30 92 30 93 30 30 00 30 00 37 00 38 00" +
                "18 00 05 96 00 00 00 17 53 77 6D 53 90 52 97 EE" +
                "68 0C 77 A9 5C 4B 62 0C 77 8F 79 F6 5C 0C 77 03" +
                "68 28 67 0C 77 FC 57 89 73 0C 77 71 67 AC 4E FD" +
                "90 43 53 49 84 0C 77 01 00 92 00 05 00 35 00 08" +
                "00 2C 00 2C 00 DB 30 C3 30 AB 30 A4 30 C9 30 A6" +
                "30 A2 30 AA 30 E2 30 EA 30 B1 30 F3 30 A4 30 EF" +
                "30 C6 30 B1 30 F3 30 D5 30 AF 30 B7 30 DE 30 B1" +
                "30 F3 30 C8 30 C1 30 AE 30 B1 30 F3 30 B5 30 A4" +
                "30 BF 30 DE 30 B1 30 F3 30 C8 30 A6 30 AD 30 E7" +
                "30 A6 30 C8 30 C1 30 D0 30 B1 30 F3 30 00 00 00" +
                "00 03 00 06 00 03 00 03 00 0C 00 06 00 03 00 11" +
                "00 09 00 03 00 17                              " +

                //Offset=0x00002636(9782) recno=72 sid=0x003C size=0x010D(269)
                "                  3C 00 0D 01 00 0C 00 03 00 1C" +
                "00 0F 00 03 00 22 00 12 00 03 00 28 00 15 00 03" +
                "00 32 00 05 10 00 00 00 42 30 44 30 46 30 48 30" +
                "4A 30 4B 30 4D 30 4F 30 51 30 53 30 55 30 57 30" +
                "59 30 5B 30 5D 30 5F 30 61 30 64 30 66 30 68 30" +
                "6A 30 6B 30 6C 30 6D 30 6E 30 6F 30 72 30 75 30" +
                "78 30 7B 30 7E 30 7F 30 80 30 81 30 82 30 84 30" +
                "86 30 88 30 89 30 8A 30 8B 30 8C 30 8D 30 8F 30" +
                "92 30 93 30 30 00 30 00 37 00 39 00 01 00 0C 00" +
                "05 00 35 00 00 00 00 00 00 00 00 00 32 00 05 10" +
                "00 00 00 42 30 44 30 46 30 48 30 4A 30 4B 30 4D" +
                "30 4F 30 51 30 53 30 55 30 57 30 59 30 5B 30 5D" +
                "30 5F 30 61 30 64 30 66 30 68 30 6A 30 6B 30 6C" +
                "30 6D 30 6E 30 6F 30 72 30 75 30 78 30 7B 30 7E" +
                "30 7F 30 80 30 81 30 82 30 84 30 86 30 88 30 89" +
                "30 8A 30 8B 30 8C 30 8D 30 8F 30 92 30 93 30 30" +
                "00 30 00 38 00 30 00 01 00 0C 00 05 00 35 00 00" +
                "00 00 00 00 00 4B 30                           ";


        /**
         * deep comparison of two SST records
         */
        private static void assertRecordEquals(SSTRecord expected, SSTRecord actual)
        {
            Assert.AreEqual(expected.NumStrings, actual.NumStrings, "number of strings");
            Assert.AreEqual(expected.NumUniqueStrings, actual.NumUniqueStrings, "number of unique strings");
            Assert.AreEqual(expected.CountStrings, actual.CountStrings, "count of strings");
            for (int k = 0; k < expected.CountStrings; k++)
            {
                UnicodeString us1 = expected.GetString(k);
                UnicodeString us2 = actual.GetString(k);

                Assert.IsTrue(us1.Equals(us2), "String at idx=" + k);
            }
        }

        [Test]
        public void Test50779_1()
        {
            byte[] bytes = HexRead.ReadFromString(data_50779_1);

            RecordInputStream in1 = TestcaseRecordInputStream.Create(bytes);
            Assert.AreEqual(SSTRecord.sid, in1.Sid);
            SSTRecord src = new SSTRecord(in1);
            Assert.AreEqual(81, src.NumStrings);

            byte[] Serialized = src.Serialize();

            in1 = TestcaseRecordInputStream.Create(Serialized);
            Assert.AreEqual(SSTRecord.sid, in1.Sid);
            SSTRecord dst = new SSTRecord(in1);
            Assert.AreEqual(81, dst.NumStrings);

            assertRecordEquals(src, dst);
        }
        [Test]
        public void Test50779_2()
        {
            byte[] bytes = HexRead.ReadFromString(data_50779_2);

            RecordInputStream in1 = TestcaseRecordInputStream.Create(bytes);
            Assert.AreEqual(SSTRecord.sid, in1.Sid);
            SSTRecord src = new SSTRecord(in1);
            Assert.AreEqual(81, src.NumStrings);

            byte[] Serialized = src.Serialize();

            in1 = TestcaseRecordInputStream.Create(Serialized);
            Assert.AreEqual(SSTRecord.sid, in1.Sid);
            SSTRecord dst = new SSTRecord(in1);
            Assert.AreEqual(81, dst.NumStrings);

            assertRecordEquals(src, dst);
        }

    }

}