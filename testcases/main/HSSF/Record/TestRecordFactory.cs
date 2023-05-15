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
    using System;
    using System.IO;
    using System.Collections;
    using NPOI.HSSF.Record;
    using NUnit.Framework;
    using NPOI.Util;
    using NPOI.POIFS.FileSystem;
    using System.Collections.Generic;

    /**
     * Tests the record factory
     * @author Glen Stampoultzis (glens at apache.org)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Csaba Nagy (ncsaba at yahoo dot com)
     */
    [TestFixture]
    public class TestRecordFactory
    {
        /// <summary>
        ///  Some of the tests are depending on the american culture.
        /// </summary>
        [SetUp]
        public void InitializeCultere()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }

        /**
         * TEST NAME:  Test Basic Record Construction <P>
         * OBJECTIVE:  Test that the RecordFactory given the required parameters for know
         *             record types can construct the proper record w/values.<P>
         * SUCCESS:    Record factory creates the records with the expected values.<P>
         * FAILURE:    The wrong records are creates or contain the wrong values <P>
         *
         */
        [Test]
        public void TestBasicRecordConstruction()
        {
            short recType = BOFRecord.sid;
            byte[] data = {
            0, 6, 5, 0, unchecked((byte)-2), 28, unchecked((byte)-51), 7, unchecked((byte)-55), 64, 0, 0, 6, 1, 0, 0
        };
            Record[] record = RecordFactory.CreateRecord(TestcaseRecordInputStream.Create(recType, data));

            Assert.AreEqual(typeof(BOFRecord).Name,
                         record[0].GetType().Name);
            BOFRecord bofRecord = (BOFRecord)record[0];

            Assert.AreEqual(7422, bofRecord.Build);
            Assert.AreEqual(1997, bofRecord.BuildYear);
            Assert.AreEqual(16585, bofRecord.HistoryBitMask);
            Assert.AreEqual(20, bofRecord.RecordSize);
            Assert.AreEqual(262, bofRecord.RequiredVersion);
            Assert.AreEqual(2057, bofRecord.Sid);
            Assert.AreEqual(BOFRecordType.Workbook, bofRecord.Type);
            Assert.AreEqual(1536, bofRecord.Version);
            recType = MMSRecord.sid;
            //size = 2;
            data = new byte[]
            {
                0, 0
            };
            record = RecordFactory.CreateRecord(TestcaseRecordInputStream.Create(recType, data));
            Assert.AreEqual(typeof(MMSRecord).Name,
                         record[0].GetType().Name);
            MMSRecord mmsRecord = (MMSRecord)record[0];

            Assert.AreEqual(0, mmsRecord.AddMenuCount);
            Assert.AreEqual(0, mmsRecord.DelMenuCount);
            Assert.AreEqual(6, mmsRecord.RecordSize);
            Assert.AreEqual(193, mmsRecord.Sid);
        }

        /**
         * TEST NAME:  Test Special Record Construction <P>
         * OBJECTIVE:  Test that the RecordFactory given the required parameters for
         *             constructing a RKRecord will return a NumberRecord.<P>
         * SUCCESS:    Record factory creates the Number record with the expected values.<P>
         * FAILURE:    The wrong records are created or contain the wrong values <P>
         *
         */
        [Test]
        public void TestSpecial()
        {
            short recType = RKRecord.sid;
            byte[] data = {
            0, 0, 0, 0, 21, 0, 0, 0, 0, 0
            };
            Record[] record = RecordFactory.CreateRecord(TestcaseRecordInputStream.Create(recType, data));

            Assert.AreEqual(typeof(NumberRecord).Name,
                         record[0].GetType().Name);
            NumberRecord numberRecord = (NumberRecord)record[0];

            Assert.AreEqual(0, numberRecord.Column);
            Assert.AreEqual(18, numberRecord.RecordSize);
            Assert.AreEqual(0, numberRecord.Row);
            Assert.AreEqual(515, numberRecord.Sid);
            Assert.AreEqual(0.0, numberRecord.Value, 0.001);
            Assert.AreEqual(21, numberRecord.XFIndex);
        }

        /**
         * TEST NAME:  Test Creating ContinueRecords After Unknown Records From An Stream <P>
         * OBJECTIVE:  Test that the RecordFactory given an Stream
         *             constructs the expected array of records.<P>
         * SUCCESS:    Record factory creates the expected records.<P>
         * FAILURE:    The wrong records are created or contain the wrong values <P>
         *
         */
        [Test]
        public void TestContinuedUnknownRecord()
        {
            byte[] data = {
            0, unchecked((byte)-1), 0, 0, // an unknown record with 0 length
            0x3C , 0, 3, 0, 1, 2, 3, // a continuation record with 3 bytes of data
            0x3C , 0, 1, 0, 4 // one more continuation record with 1 byte of data
            };

            MemoryStream bois = new MemoryStream(data);
            Record[] records = (Record[])
              RecordFactory.CreateRecords(bois).ToArray();
            Assert.AreEqual(3, records.Length, "Created record count");
            Assert.AreEqual(
                         typeof(UnknownRecord).Name,
                         records[0].GetType().Name, "1st record's type");
            Assert.AreEqual((short)-256, records[0].Sid, "1st record's sid");
            Assert.AreEqual(typeof(ContinueRecord).Name,
                         records[1].GetType().Name, "2nd record's type");
            ContinueRecord record = (ContinueRecord)records[1];
            Assert.AreEqual(0x3C, record.Sid, "2nd record's sid");
            Assert.AreEqual(1, record.Data[0], "1st data byte");
            Assert.AreEqual(2, record.Data[1], "2nd data byte");
            Assert.AreEqual(3, record.Data[2], "3rd data byte");
            Assert.AreEqual(
                         typeof(ContinueRecord).Name,
                         records[2].GetType().Name, "3rd record's type");
            record = (ContinueRecord)records[2];
            Assert.AreEqual(0x3C, record.Sid, "3nd record's sid");
            Assert.AreEqual(4, record.Data[0], "4th data byte");
        }

        /**
         * Drawing records have a very strange continue behaviour.
         * There can actually be OBJ records mixed between the continues.
         * Record factory must preserve this structure when Reading records.
         */
        [Test]
        public void TestMixedContinue()
        {
            /**
             *  Taken from a real test sample file 39512.xls. See Bug 39512 for details.
             */
            String dump =
                    //OBJ
                    "5D 00 48 00 15 00 12 00 0C 00 3C 00 11 00 A0 2E 03 01 CC 42 " +
                    "CF 00 00 00 00 00 0A 00 0C 00 00 00 00 00 00 00 00 00 00 00 " +
                    "03 00 0B 00 06 00 28 01 03 01 00 00 12 00 08 00 00 00 00 00 " +
                    "00 00 03 00 11 00 04 00 3D 00 00 00 00 00 00 00 " +
                    //MSODRAWING
                    "EC 00 08 00 00 00 0D F0 00 00 00 00 " +
                    //TXO (and 2 trailing CONTINUE records)
                    //"B6 01 12 00 22 02 00 00 00 00 00 00 00 00 10 00 10 00 00 00 " +
                    //"00 00 3C 00 21 00 01 4F 00 70 00 74 00 69 00 6F 00 6E 00 20 " +
                    //"00 42 00 75 00 74 00 74 00 6F 00 6E 00 20 00 33 00 39 00 3C " +
                    //"00 10 00 00 00 05 00 00 00 00 00 10 00 00 00 00 00 00 00 " +
                    "B6 01 12 00 22 02 00 00 00 00 00 00 00 00 10 00 10 00 00 00 00 00 " +
                    "3C 00 11 00 00 4F 70 74 69 6F 6E 20 42 75 74 74 6F 6E 20 33 39 " +
                    "3C 00 10 00 00 00 05 00 00 00 00 00 10 00 00 00 00 00 00 00 " +
                    // another CONTINUE
                    "3C 00 7E 00 0F 00 04 F0 7E 00 00 00 92 0C 0A F0 08 00 00 00 " +
                    "3D 04 00 00 00 0A 00 00 A3 00 0B F0 3C 00 00 00 7F 00 00 01 " +
                    "00 01 80 00 8C 01 03 01 85 00 01 00 00 00 8B 00 02 00 00 00 " +
                    "BF 00 08 00 1A 00 7F 01 29 00 29 00 81 01 41 00 00 08 BF 01 " +
                    "00 00 10 00 C0 01 40 00 00 08 FF 01 00 00 08 00 00 00 10 F0 " +
                    "12 00 00 00 02 00 02 00 A0 03 18 00 B5 00 04 00 30 02 1A 00 " +
                    "00 00 00 00 11 F0 00 00 00 00 " +
                    //OBJ
                    "5D 00 48 00 15 00 12 00 0C 00 3D 00 11 00 8C 01 03 01 C8 59 CF 00 00 " +
                    "00 00 00 0A 00 0C 00 00 00 00 00 00 00 00 00 00 00 03 00 0B 00 06 00 " +
                    "7C 16 03 01 00 00 12 00 08 00 00 00 00 00 00 00 03 00 11 00 04 00 01 " +
                    "00 00 00 00 00 00 00";
            byte[] data = HexRead.ReadFromString(dump);

            IList records = RecordFactory.CreateRecords(new MemoryStream(data));
            Assert.AreEqual(5, records.Count);
            Assert.IsTrue(records[0] is ObjRecord);
            Assert.IsTrue(records[1] is DrawingRecord);
            Assert.IsTrue(records[2] is TextObjectRecord);
            Assert.IsTrue(records[3] is ContinueRecord);
            Assert.IsTrue(records[4] is ObjRecord);

            //Serialize and verify that the Serialized data is1 the same as the original
            MemoryStream out1 = new MemoryStream();
            for (IEnumerator it = records.GetEnumerator(); it.MoveNext();)
            {
                Record rec = (Record)it.Current;
                byte[] serialdata = rec.Serialize();
                out1.Write(serialdata, 0, serialdata.Length);
            }

            byte[] ser = out1.ToArray();
            Assert.AreEqual(data.Length, ser.Length);
            Assert.IsTrue(Arrays.Equals(data, ser));
        }
        [Test]
        public void TestNonZeroPadding_bug46987()
        {
            Record[] recs = {
                new BOFRecord(),
                new WriteAccessRecord(), // need *something* between BOF and EOF
			    EOFRecord.instance,
                BOFRecord.CreateSheetBOF(),
                EOFRecord.instance,
            };
            MemoryStream baos = new MemoryStream();
            for (int i = 0; i < recs.Length; i++)
            {
                byte[] data = recs[i].Serialize();
                baos.Write(data, 0, data.Length);
            }
            //simulate the bad padding at the end of the workbook stream in attachment 23483 of bug 46987
            baos.WriteByte(0x00);
            baos.WriteByte(0x11);
            baos.WriteByte(0x00);
            baos.WriteByte(0x02);
            for (int i = 0; i < 192; i++)
            {
                baos.WriteByte(0x00);
            }


            POIFSFileSystem fs = new POIFSFileSystem();
            Stream is1;
            fs.CreateDocument(new MemoryStream(baos.ToArray()), "dummy");
            is1 = fs.Root.CreatePOIFSDocumentReader("dummy");


            List<Record> outRecs;
            try
            {
                outRecs = RecordFactory.CreateRecords(is1);
            }
            catch (Exception e)
            {
                if (e.Message.Equals("Buffer underrun - requested 512 bytes but 192 was available"))
                {
                    throw new AssertionException("Identified bug 46987");
                }
                throw;
            }
            Assert.AreEqual(5, outRecs.Count);
        }
        [Test]
        //public void TestNonZeroPadding_bug46987()
        public void TestNPOIBug6177()
        {
            //Record[] recs = {
            //    new BOFRecord(),
            //    EOFRecord.instance,
            //    BOFRecord.CreateSheetBOF(),
            //    EOFRecord.instance,
            //};
            //MemoryStream baos = new MemoryStream();
            //for (int i = 0; i < recs.Length; i++)
            //{
            //    baos.Write(recs[i].Serialize(),0, recs[i].RecordSize);
            //}
            ////simulate the bad padding at the end of the workbook stream in attachment 23483 of bug 46987
            //baos.WriteByte(0x00);
            //baos.WriteByte(0x11);
            //baos.WriteByte(0x00);
            //baos.WriteByte(0x02);
            //for (int i = 0; i < 192; i++)
            //{
            //    baos.WriteByte(0x00);
            //}


            //POIFSFileSystem fs = new POIFSFileSystem();
            //Stream is1;
            //fs.CreateDocument(new MemoryStream(baos.ToArray()), "dummy");
            //is1 = fs.Root.CreatePOIFSDocumentReader("dummy");

            //List<Record> outRecs;
            //try
            //{
            //    outRecs = RecordFactory.CreateRecords(is1);
            //}
            //catch (Exception e)
            //{
            //    if (e.Message.Equals("Buffer underrun - requested 512 bytes but 192 was available"))
            //    {
            //        throw new AssertionException("Identified bug 46987");
            //    }
            //    throw e;
            //}
            //Assert.AreEqual(4, outRecs.Count);
            string sampleFileName = "FW 8.6 Table Relationship2.xls";
            HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }

    }
}
