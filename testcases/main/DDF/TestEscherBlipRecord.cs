
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

    /**
     * Test Read/Serialize of escher blip records
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestEscherBlipRecord
    {
        static POIDataSamples _samples = POIDataSamples.GetDDFInstance();


        //Test Reading/serializing of a PNG blip
        [Test]
        public void TestReadPNG()
        {
            //provided in bug-44886
            byte[] data = _samples.ReadFile("Container.dat");

            EscherContainerRecord record = new EscherContainerRecord();
            record.FillFields(data, 0, new DefaultEscherRecordFactory());
            EscherContainerRecord bstore = (EscherContainerRecord)record.ChildRecords[1];
            EscherBSERecord bse1 = (EscherBSERecord)bstore.ChildRecords[0];
            Assert.AreEqual(EscherBSERecord.BT_PNG, bse1.BlipTypeWin32);
            Assert.AreEqual(EscherBSERecord.BT_PNG, bse1.BlipTypeMacOS);
            Assert.IsTrue(Arrays.Equals(new byte[]{
            0x65, 0x07, 0x4A, (byte)0x8D, 0x3E, 0x42, (byte)0x8B, (byte)0xAC,
            0x1D, (byte)0x89, 0x35, 0x4F, 0x48, (byte)0xFA, 0x37, (byte)0xC2
        }, bse1.UID));
            Assert.AreEqual(255, bse1.Tag);
            Assert.AreEqual(32308, bse1.Size);

            EscherBitmapBlip blip1 = (EscherBitmapBlip)bse1.BlipRecord;
            Assert.AreEqual(0x6E00, blip1.Options);
            Assert.AreEqual(EscherBitmapBlip.RECORD_ID_PNG, blip1.RecordId);
            Assert.IsTrue(Arrays.Equals(new byte[]{
            0x65, 0x07, 0x4A, (byte)0x8D, 0x3E, 0x42, (byte)0x8B, (byte)0xAC,
            0x1D, (byte)0x89, 0x35, 0x4F, 0x48, (byte)0xFA, 0x37, (byte)0xC2
        }, blip1.UID));

            //Serialize and Read again
            byte[] ser = bse1.Serialize();
            EscherBSERecord bse2 = new EscherBSERecord();
            bse2.FillFields(ser, 0, new DefaultEscherRecordFactory());
            Assert.AreEqual(bse1.RecordId, bse2.RecordId);
            Assert.AreEqual(bse1.BlipTypeWin32, bse2.BlipTypeWin32);
            Assert.AreEqual(bse1.BlipTypeMacOS, bse2.BlipTypeMacOS);
            Assert.IsTrue(Arrays.Equals(bse1.UID, bse2.UID));
            Assert.AreEqual(bse1.Tag, bse2.Tag);
            Assert.AreEqual(bse1.Size, bse2.Size);

            EscherBitmapBlip blip2 = (EscherBitmapBlip)bse1.BlipRecord;
            Assert.AreEqual(blip1.Options, blip2.Options);
            Assert.AreEqual(blip1.RecordId, blip2.RecordId);
            Assert.AreEqual(blip1.UID, blip2.UID);

            Assert.IsTrue(Arrays.Equals(blip1.PictureData, blip1.PictureData));
        }

        //Test Reading/serializing of a PICT metafile
        [Test]
        public void TestReadPICT()
        {
            //provided in bug-44886
            byte[] data = _samples.ReadFile("Container.dat");

            EscherContainerRecord record = new EscherContainerRecord();
            record.FillFields(data, 0, new DefaultEscherRecordFactory());
            EscherContainerRecord bstore = (EscherContainerRecord)record.ChildRecords[1];
            EscherBSERecord bse1 = (EscherBSERecord)bstore.ChildRecords[1];
            //System.out.println(bse1);
            Assert.AreEqual(EscherBSERecord.BT_WMF, bse1.BlipTypeWin32);
            Assert.AreEqual(EscherBSERecord.BT_PICT, bse1.BlipTypeMacOS);
            Assert.IsTrue(Arrays.Equals(new byte[]{
            (byte)0xC7, 0x15, 0x69, 0x2D, (byte)0xE5, (byte)0x89, (byte)0xA3, 0x6F,
            0x66, 0x03, (byte)0xD6, 0x24, (byte)0xF7, (byte)0xDB, 0x1D, 0x13
        }, bse1.UID));
            Assert.AreEqual(255, bse1.Tag);
            Assert.AreEqual(1133, bse1.Size);

            EscherMetafileBlip blip1 = (EscherMetafileBlip)bse1.BlipRecord;
            Assert.AreEqual(0x5430, blip1.Options);
            Assert.AreEqual(EscherMetafileBlip.RECORD_ID_PICT, blip1.RecordId);
            Assert.IsTrue(Arrays.Equals(new byte[]{
            0x57, 0x32, 0x7B, (byte)0x91, 0x23, 0x5D, (byte)0xDB, 0x36,
            0x7A, (byte)0xDB, (byte)0xFF, 0x17, (byte)0xFE, (byte)0xF3, (byte)0xA7, 0x05
        }, blip1.UID));
            Assert.IsTrue(Arrays.Equals(new byte[]{
            (byte)0xC7, 0x15, 0x69, 0x2D, (byte)0xE5, (byte)0x89, (byte)0xA3, 0x6F,
            0x66, 0x03, (byte)0xD6, 0x24, (byte)0xF7, (byte)0xDB, 0x1D, 0x13
        }, blip1.PrimaryUID));

            //Serialize and Read again
            byte[] ser = bse1.Serialize();
            EscherBSERecord bse2 = new EscherBSERecord();
            bse2.FillFields(ser, 0, new DefaultEscherRecordFactory());
            Assert.AreEqual(bse1.RecordId, bse2.RecordId);
            Assert.AreEqual(bse1.Options, bse2.Options);
            Assert.AreEqual(bse1.BlipTypeWin32, bse2.BlipTypeWin32);
            Assert.AreEqual(bse1.BlipTypeMacOS, bse2.BlipTypeMacOS);
            Assert.IsTrue(Arrays.Equals(bse1.UID, bse2.UID));
            Assert.AreEqual(bse1.Tag, bse2.Tag);
            Assert.AreEqual(bse1.Size, bse2.Size);

            EscherMetafileBlip blip2 = (EscherMetafileBlip)bse1.BlipRecord;
            Assert.AreEqual(blip1.Options, blip2.Options);
            Assert.AreEqual(blip1.RecordId, blip2.RecordId);
            Assert.AreEqual(blip1.UID, blip2.UID);
            Assert.AreEqual(blip1.PrimaryUID, blip2.PrimaryUID);

            Assert.IsTrue(Arrays.Equals(blip1.PictureData, blip1.PictureData));
        }

        //integral Test: check that the Read-Write-Read round trip is consistent
        [Test]
        public void TestContainer()
        {
            byte[] data = _samples.ReadFile("Container.dat");

            EscherContainerRecord record = new EscherContainerRecord();
            record.FillFields(data, 0, new DefaultEscherRecordFactory());

            byte[] ser = record.Serialize();
            Assert.IsTrue(Arrays.Equals(data, ser));
        }

        /**
     * The test data was created from pl031405.xls attached to Bugzilla #47143
     */
        [Test]
        public void Test47143()
        {
            byte[] data = _samples.ReadFile("47143.dat");
            EscherBSERecord bse = new EscherBSERecord();
            bse.FillFields(data, 0, new DefaultEscherRecordFactory());
            bse.ToString(); //assert that toString() works
            Assert.IsTrue(bse.BlipRecord is EscherMetafileBlip);

            EscherMetafileBlip blip = (EscherMetafileBlip)bse.BlipRecord;
            blip.ToString(); //assert that toString() works
            byte[] remaining = blip.RemainingData;
            Assert.IsNotNull(remaining);

            byte[] ser = bse.Serialize();  //serialize and assert against the source data
            Assert.IsTrue(Arrays.Equals(data, ser));
        }
    }
}