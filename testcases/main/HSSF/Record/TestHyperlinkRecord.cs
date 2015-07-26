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
    using System.Web;
    using System.IO;
    using NPOI.Util;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Util;



    using NUnit.Framework;

    /**
     * Test HyperlinkRecord
     *
     * @author Nick Burch
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestHyperlinkRecord
    {
        /// <summary>
        ///  Some of the tests are depending on the american culture.
        /// </summary>
        [SetUp]
        public void InitializeCultere()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }


        //link to http://www.lakings.com/
        byte[] data1 = { 0x02, 0x00,    //First row of the hyperlink
                     0x02, 0x00,    //Last row of the hyperlink
                     0x00, 0x00,    //First column of the hyperlink
                     0x00, 0x00,    //Last column of the hyperlink

                     //16-byte GUID. Seems to be always the same. Does not depend on the hyperlink type
                     (byte)0xD0, (byte)0xC9, (byte)0xEA, 0x79, (byte)0xF9, (byte)0xBA, (byte)0xCE, 0x11,
                     (byte)0x8C, (byte)0x82, 0x00, (byte)0xAA, 0x00, 0x4B, (byte)0xA9, 0x0B,

                    0x02, 0x00, 0x00, 0x00, //integer, always 2

                    // flags. Define the type of the hyperlink:
                    // HyperlinkRecord.HLINK_URL | HyperlinkRecord.HLINK_ABS | HyperlinkRecord.HLINK_LABEL
                    0x17, 0x00, 0x00, 0x00,

                    0x08, 0x00, 0x00, 0x00, //length of the label including the trailing '\0'

                    //label:
                    0x4D, 0x00, 0x79, 0x00, 0x20, 0x00, 0x4C, 0x00, 0x69, 0x00, 0x6E, 0x00, 0x6B, 0x00, 0x00, 0x00,

                    //16-byte link moniker: HyperlinkRecord.URL_MONIKER
                    (byte)0xE0, (byte)0xC9, (byte)0xEA, 0x79, (byte)0xF9, (byte)0xBA, (byte)0xCE,  0x11,
                    (byte)0x8C, (byte)0x82, 0x00, (byte)0xAA, 0x00, 0x4B, (byte)0xA9, 0x0B,

                    //count of bytes in the address including the tail
                    0x48, 0x00, 0x00, 0x00, //integer

                    //the actual link, terminated by '\u0000'
                    0x68, 0x00, 0x74, 0x00, 0x74, 0x00, 0x70, 0x00, 0x3A, 0x00, 0x2F, 0x00,
                    0x2F, 0x00, 0x77, 0x00, 0x77, 0x00, 0x77, 0x00, 0x2E, 0x00, 0x6C, 0x00,
                    0x61, 0x00, 0x6B, 0x00, 0x69, 0x00, 0x6E, 0x00, 0x67, 0x00, 0x73, 0x00,
                    0x2E, 0x00, 0x63, 0x00, 0x6F, 0x00, 0x6D, 0x00, 0x2F, 0x00, 0x00, 0x00,

                    //standard 24-byte tail of a URL link. Seems to always be the same for all URL HLINKs
                    0x79, 0x58, (byte)0x81, (byte)0xF4, 0x3B, 0x1D, 0x7F, 0x48, (byte)0xAF, 0x2C,
                    (byte)0x82, 0x5D, (byte)0xC4, (byte)0x85, 0x27, 0x63, 0x00, 0x00, 0x00,
                    0x00, (byte)0xA5, (byte)0xAB, 0x00, 0x00};

        //link to a file in the current directory: link1.xls
        byte[] data2 =  {0x00, 0x00,
                     0x00, 0x00,
                     0x00, 0x00,
                     0x00, 0x00,
                     //16-bit GUID. Seems to be always the same. Does not depend on the hyperlink type
                     (byte)0xD0, (byte)0xC9, (byte)0xEA, 0x79, (byte)0xF9, (byte)0xBA, (byte)0xCE, 0x11,
                     (byte)0x8C, (byte)0x82, 0x00, (byte)0xAA, 0x00, 0x4B, (byte)0xA9, 0x0B,

                     0x02, 0x00, 0x00, 0x00,    //integer, always 2

                     0x15, 0x00, 0x00, 0x00,    //options: HyperlinkRecord.HLINK_URL | HyperlinkRecord.HLINK_LABEL

                     0x05, 0x00, 0x00, 0x00,    //length of the label
                     //label
                     0x66, 0x00, 0x69, 0x00, 0x6C, 0x00, 0x65, 0x00, 0x00, 0x00,

                     //16-byte link moniker: HyperlinkRecord.FILE_MONIKER
                     0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, (byte)0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,

                     0x00, 0x00,    //level
                     0x0A, 0x00, 0x00, 0x00,    //length of the path )

                     //path to the file (plain ISO-8859 bytes, NOT UTF-16LE!)
                     0x6C, 0x69, 0x6E, 0x6B, 0x31, 0x2E, 0x78, 0x6C, 0x73, 0x00,

                     //standard 28-byte tail of a file link
                     (byte)0xFF, (byte)0xFF, (byte)0xAD, (byte)0xDE, 0x00, 0x00, 0x00, 0x00,
                     0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                     0x00, 0x00, 0x00, 0x00, 
                     
                     0x00, 0x00, 0x00, 0x00   
                    };

        // mailto:ebgans@mail.ru?subject=Hello,%20Ebgans!
        byte[] data3 = {0x01, 0x00,
                    0x01, 0x00,
                    0x00, 0x00,
                    0x00, 0x00,

                    //16-bit GUID. Seems to be always the same. Does not depend on the hyperlink type
                    (byte)0xD0, (byte)0xC9, (byte)0xEA, 0x79, (byte)0xF9, (byte)0xBA, (byte)0xCE, 0x11,
                    (byte)0x8C, (byte)0x82, 0x00, (byte)0xAA, 0x00, 0x4B, (byte)0xA9, 0x0B,

                    0x02, 0x00, 0x00, 0x00, //integer, always 2

                    0x17, 0x00, 0x00, 0x00,  //options: HyperlinkRecord.HLINK_URL | HyperlinkRecord.HLINK_ABS | HyperlinkRecord.HLINK_LABEL

                    0x06, 0x00, 0x00, 0x00,     //length of the label
                    0x65, 0x00, 0x6D, 0x00, 0x61, 0x00, 0x69, 0x00, 0x6C, 0x00, 0x00, 0x00, //label

                    //16-byte link moniker: HyperlinkRecord.URL_MONIKER
                    (byte)0xE0, (byte)0xC9, (byte)0xEA, 0x79, (byte)0xF9, (byte)0xBA, (byte)0xCE, 0x11,
                    (byte)0x8C, (byte)0x82, 0x00, (byte)0xAA, 0x00, 0x4B, (byte)0xA9, 0x0B,

                    //length of the address including the tail.
                    0x76, 0x00, 0x00, 0x00,

                    //the address is terminated by '\u0000'
                    0x6D, 0x00, 0x61, 0x00, 0x69, 0x00, 0x6C, 0x00, 0x74, 0x00, 0x6F, 0x00,
                    0x3A, 0x00, 0x65, 0x00, 0x62, 0x00, 0x67, 0x00, 0x61, 0x00, 0x6E, 0x00,
                    0x73, 0x00, 0x40, 0x00, 0x6D, 0x00, 0x61, 0x00, 0x69, 0x00, 0x6C, 0x00,
                    0x2E, 0x00, 0x72, 0x00, 0x75, 0x00, 0x3F, 0x00, 0x73, 0x00, 0x75, 0x00,
                    0x62, 0x00, 0x6A, 0x00, 0x65, 0x00, 0x63, 0x00, 0x74, 0x00, 0x3D, 0x00,
                    0x48, 0x00, 0x65, 0x00, 0x6C, 0x00, 0x6C, 0x00, 0x6F, 0x00, 0x2C, 0x00,
                    0x25, 0x00, 0x32, 0x00, 0x30, 0x00, 0x45, 0x00, 0x62, 0x00, 0x67, 0x00,
                    0x61, 0x00, 0x6E, 0x00, 0x73, 0x00, 0x21, 0x00, 0x00, 0x00,

                    //standard 24-byte tail of a URL link
                    0x79, 0x58, (byte)0x81, (byte)0xF4, 0x3B, 0x1D, 0x7F, 0x48, (byte)0xAF, (byte)0x2C,
                    (byte)0x82, 0x5D, (byte)0xC4, (byte)0x85, 0x27, 0x63, 0x00, 0x00, 0x00,
                    0x00, (byte)0xA5, (byte)0xAB, 0x00, 0x00
    };

        //link to a place in worksheet: Sheet1!A1
        byte[] data4 = {0x03, 0x00,
                    0x03, 0x00,
                    0x00, 0x00,
                    0x00, 0x00,

                    //16-bit GUID. Seems to be always the same. Does not depend on the hyperlink type
                    (byte)0xD0, (byte)0xC9, (byte)0xEA, 0x79, (byte)0xF9, (byte)0xBA, (byte)0xCE, 0x11,
                    (byte)0x8C, (byte)0x82, 0x00, (byte)0xAA, 0x00, 0x4B, (byte)0xA9, 0x0B,

                    0x02, 0x00, 0x00, 0x00, //integer, always 2

                    0x1C, 0x00, 0x00, 0x00, //flags: HyperlinkRecord.HLINK_LABEL | HyperlinkRecord.HLINK_PLACE

                    0x06, 0x00, 0x00, 0x00, //length of the label

                    0x70, 0x00, 0x6C, 0x00, 0x61, 0x00, 0x63, 0x00, 0x65, 0x00, 0x00, 0x00, //label

                    0x0A, 0x00, 0x00, 0x00, //length of the document link including trailing zero

                    //link: Sheet1!A1
                    0x53, 0x00, 0x68, 0x00, 0x65, 0x00, 0x65, 0x00, 0x74, 0x00, 0x31, 0x00, 0x21,
                    0x00, 0x41, 0x00, 0x31, 0x00, 0x00, 0x00};

        private static byte[] dataLinkToWorkbook = HexRead.ReadFromString("01 00 01 00 01 00 01 00 " +
        "D0 C9 EA 79 F9 BA CE 11 8C 82 00 AA 00 4B A9 0B " +
        "02 00 00 00 " +
        "1D 00 00 00 " + // options: LABEL | PLACE | FILE_OR_URL
            // label: "My Label"
        "09 00 00 00 " +
        "4D 00 79 00 20 00 4C 00 61 00 62 00 65 00 6C 00 00 00 " +
        "03 03 00 00 00 00 00 00 C0 00 00 00 00 00 00 46 " + // file GUID
        "00 00 " + // file options
            // shortFileName: "YEARFR~1.XLS"
        "0D 00 00 00 " +
        "59 45 41 52 46 52 7E 31 2E 58 4C 53 00 " +
            // FILE_TAIL - unknown byte sequence
        "FF FF AD DE 00 00 00 00 " +
        "00 00 00 00 00 00 00 00 " +
        "00 00 00 00 00 00 00 00 " +
            // field len, char data len
        "2E 00 00 00 " +
        "28 00 00 00 " +
        "03 00 " + // unknown ushort
            // _address: "yearfracExamples.xls"
        "79 00 65 00 61 00 72 00 66 00 72 00 61 00 63 00 " +
        "45 00 78 00 61 00 6D 00 70 00 6C 00 65 00 73 00 " +
        "2E 00 78 00 6C 00 73 00 " +
            // textMark: "Sheet1!B6"
        "0A 00 00 00 " +
        "53 00 68 00 65 00 65 00 74 00 31 00 21 00 42 00 " +
        "36 00 00 00");

        private static byte[] dataTargetFrame = HexRead.ReadFromString("0E 00 0E 00 00 00 00 00 " +
                "D0 C9 EA 79 F9 BA CE 11  8C 82 00 AA 00 4B A9 0B " +
                "02 00 00 00 " +
                "83 00 00 00 " + // options: TARGET_FRAME | ABS | FILE_OR_URL
            // targetFrame: "_blank"
                "07 00 00 00 " +
                "5F 00 62 00 6C 00 61 00 6E 00 6B 00 00 00 " +
            // url GUID
                "E0 C9 EA 79 F9 BA CE 11 8C 82 00 AA 00 4B A9 0B " +
            // address: "http://www.regnow.com/softsell/nph-softsell.cgi?currency=USD&item=7924-37"
                "94 00 00 00 " +
                "68 00 74 00 74 00 70 00 3A 00 2F 00 2F 00 77 00 " +
                "77 00 77 00 2E 00 72 00 65 00 67 00 6E 00 6F 00 " +
                "77 00 2E 00 63 00 6F 00 6D 00 2F 00 73 00 6F 00 " +
                "66 00 74 00 73 00 65 00 6C 00 6C 00 2F 00 6E 00 " +
                "70 00 68 00 2D 00 73 00 6F 00 66 00 74 00 73 00 " +
                "65 00 6C 00 6C 00 2E 00 63 00 67 00 69 00 3F 00 " +
                "63 00 75 00 72 00 72 00 65 00 6E 00 63 00 79 00 " +
                "3D 00 55 00 53 00 44 00 26 00 69 00 74 00 65 00 " +
                "6D 00 3D 00 37 00 39 00 32 00 34 00 2D 00 33 00 " +
                "37 00 00 00");


        private static byte[] dataUNC = HexRead.ReadFromString("01 00 01 00 01 00 01 00 " +
                "D0 C9 EA 79 F9 BA CE 11 8C 82 00 AA 00 4B A9 0B " +
                "02 00 00 00 " +
                "1F 01 00 00 " + // options: UNC_PATH | LABEL | TEXT_MARK | ABS | FILE_OR_URL
                "09 00 00 00 " + // label: "My Label"
                "4D 00 79 00 20 00 6C 00 61 00 62 00 65 00 6C 00 00 00 " +
            // note - no moniker GUID
                "27 00 00 00 " +  // "\\\\MyServer\\my-share\\myDir\\PRODNAME.xls"
                "5C 00 5C 00 4D 00 79 00 53 00 65 00 72 00 76 00 " +
                "65 00 72 00 5C 00 6D 00 79 00 2D 00 73 00 68 00 " +
                "61 00 72 00 65 00 5C 00 6D 00 79 00 44 00 69 00 " +
                "72 00 5C 00 50 00 52 00 4F 00 44 00 4E 00 41 00 " +
                "4D 00 45 00 2E 00 78 00 6C 00 73 00 00 00 " +

                "0C 00 00 00 " + // textMark: PRODNAME!C2
                "50 00 52 00 4F 00 44 00 4E 00 41 00 4D 00 45 00 21 00 " +
                "43 00 32 00 00 00");


        /**
         * From Bugzilla 47498
         */
        private static byte[] data_47498 = {
                0x02, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, (byte)0xD0, (byte)0xC9,
                (byte)0xEA, 0x79, (byte)0xF9, (byte)0xBA, (byte)0xCE, 0x11, (byte)0x8C,
                (byte)0x82, 0x00, (byte)0xAA, 0x00, 0x4B, (byte)0xA9, 0x0B, 0x02, 0x00,
                0x00, 0x00, 0x15, 0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x50, 0x00,
                0x44, 0x00, 0x46, 0x00, 0x00, 0x00, (byte)0xE0, (byte)0xC9, (byte)0xEA,
                0x79, (byte)0xF9, (byte)0xBA, (byte)0xCE, (byte)0x11, (byte)0x8C, (byte)0x82,
                0x00, (byte)0xAA, 0x00, 0x4B, (byte)0xA9, 0x0B, 0x28, 0x00, 0x00, 0x00,
                0x74, 0x00, 0x65, 0x00, 0x73, 0x00, 0x74, 0x00, 0x66, 0x00, 0x6F, 0x00,
                0x6C, 0x00, 0x64, 0x00, 0x65, 0x00, 0x72, 0x00, 0x2F, 0x00, 0x74, 0x00,
                0x65, 0x00, 0x73, 0x00, 0x74, 0x00, 0x2E, 0x00, 0x50, 0x00, 0x44, 0x00,
                0x46, 0x00, 0x00, 0x00};


        private void ConfirmGUID(GUID expectedGuid, GUID actualGuid)
        {
            Assert.AreEqual(expectedGuid, actualGuid);
        }

        [Test]
        public void TestReadURLLink()
        {
            RecordInputStream is1 = TestcaseRecordInputStream.Create(HyperlinkRecord.sid, data1);
            HyperlinkRecord link = new HyperlinkRecord(is1);
            Assert.AreEqual(2, link.FirstRow);
            Assert.AreEqual(2, link.LastRow);
            Assert.AreEqual(0, link.FirstColumn);
            Assert.AreEqual(0, link.LastColumn);
            ConfirmGUID(HyperlinkRecord.STD_MONIKER, link.Guid);
            ConfirmGUID(HyperlinkRecord.URL_MONIKER, link.Moniker);
            Assert.AreEqual(2, link.LabelOptions);
            int opts = HyperlinkRecord.HLINK_URL | HyperlinkRecord.HLINK_ABS | HyperlinkRecord.HLINK_LABEL;
            Assert.AreEqual(0x17, opts);
            Assert.AreEqual(opts, link.LinkOptions);
            Assert.AreEqual(0, link.FileOptions);

            Assert.AreEqual("My Link", link.Label);
            Assert.AreEqual("http://www.lakings.com/", link.Address);
        }
        [Test]
        public void TestReadFileLink()
        {
            RecordInputStream is1 = TestcaseRecordInputStream.Create((short)HyperlinkRecord.sid, data2);
            HyperlinkRecord link = new HyperlinkRecord(is1);
            Assert.AreEqual(0, link.FirstRow);
            Assert.AreEqual(0, link.LastRow);
            Assert.AreEqual(0, link.FirstColumn);
            Assert.AreEqual(0, link.LastColumn);
            ConfirmGUID(HyperlinkRecord.STD_MONIKER, link.Guid);
            ConfirmGUID(HyperlinkRecord.FILE_MONIKER, link.Moniker);
            Assert.AreEqual(2, link.LabelOptions);
            int opts = HyperlinkRecord.HLINK_URL | HyperlinkRecord.HLINK_LABEL;
            Assert.AreEqual(0x15, opts);
            Assert.AreEqual(opts, link.LinkOptions);

            Assert.AreEqual("file", link.Label);
            Assert.AreEqual("link1.xls", link.ShortFilename);
        }
        [Test]
        public void TestReadEmailLink()
        {
            RecordInputStream is1 = TestcaseRecordInputStream.Create((short)HyperlinkRecord.sid, data3);
            HyperlinkRecord link = new HyperlinkRecord(is1);
            Assert.AreEqual(1, link.FirstRow);
            Assert.AreEqual(1, link.LastRow);
            Assert.AreEqual(0, link.FirstColumn);
            Assert.AreEqual(0, link.LastColumn);
            ConfirmGUID(HyperlinkRecord.STD_MONIKER, link.Guid);
            ConfirmGUID(HyperlinkRecord.URL_MONIKER, link.Moniker);
            Assert.AreEqual(2, link.LabelOptions);
            int opts = HyperlinkRecord.HLINK_URL | HyperlinkRecord.HLINK_ABS | HyperlinkRecord.HLINK_LABEL;
            Assert.AreEqual(0x17, opts);
            Assert.AreEqual(opts, link.LinkOptions);

            Assert.AreEqual("email", link.Label);
            Assert.AreEqual("mailto:ebgans@mail.ru?subject=Hello,%20Ebgans!", link.Address);
        }
        [Test]
        public void TestReadDocumentLink()
        {
            RecordInputStream is1 = TestcaseRecordInputStream.Create(HyperlinkRecord.sid, data4);
            HyperlinkRecord link = new HyperlinkRecord(is1);
            Assert.AreEqual(3, link.FirstRow);
            Assert.AreEqual(3, link.LastRow);
            Assert.AreEqual(0, link.FirstColumn);
            Assert.AreEqual(0, link.LastColumn);
            ConfirmGUID(HyperlinkRecord.STD_MONIKER, link.Guid);
            Assert.AreEqual(2, link.LabelOptions);
            int opts = HyperlinkRecord.HLINK_LABEL | HyperlinkRecord.HLINK_PLACE;
            Assert.AreEqual(0x1C, opts);
            Assert.AreEqual(opts, link.LinkOptions);

            Assert.AreEqual("place", link.Label);
            Assert.AreEqual("Sheet1!A1", link.TextMark);
            Assert.AreEqual("Sheet1!A1", link.Address);
        }
        private void Serialize(byte[] data)
        {
            RecordInputStream is1 = TestcaseRecordInputStream.Create(HyperlinkRecord.sid, data);
            HyperlinkRecord link = new HyperlinkRecord(is1);
            byte[] bytes1 = link.Serialize();
            is1 = TestcaseRecordInputStream.Create(bytes1);
            link = new HyperlinkRecord(is1);
            byte[] bytes2 = link.Serialize();
            Assert.AreEqual(bytes1.Length, bytes2.Length);
            Assert.IsTrue(Arrays.Equals(bytes1, bytes2));
        }
        [Test]
        public void TestSerialize()
        {
            Serialize(data1);
            Serialize(data2);
            Serialize(data3);
            Serialize(data4);
        }
        [Test]
        public void TestCreateURLRecord()
        {
            HyperlinkRecord link = new HyperlinkRecord();
            link.CreateUrlLink();
            link.FirstRow = 2;
            link.LastRow = 2;
            link.Label = "My Link";
            link.Address = "http://www.lakings.com/";

            byte[] tmp = link.Serialize();
            byte[] ser = new byte[tmp.Length - 4];
            Array.Copy(tmp, 4, ser, 0, ser.Length);
            Assert.AreEqual(data1.Length, ser.Length);
            Assert.IsTrue(Arrays.Equals(data1, ser));
        }
        [Test]
        public void TestCreateFileRecord()
        {
            HyperlinkRecord link = new HyperlinkRecord();
            link.CreateFileLink();
            link.FirstRow = 0;
            link.LastRow = 0;
            link.Label = "file";
            link.ShortFilename = "link1.xls";

            byte[] tmp = link.Serialize();
            byte[] ser = new byte[tmp.Length - 4];
            Array.Copy(tmp, 4, ser, 0, ser.Length);
            Assert.AreEqual(data2.Length, ser.Length);
            Assert.IsTrue(Arrays.Equals(data2, ser));
        }
        [Test]
        public void TestCreateDocumentRecord()
        {
            HyperlinkRecord link = new HyperlinkRecord();
            link.CreateDocumentLink();
            link.FirstRow = 3;
            link.LastRow = 3;
            link.Label = "place";
            link.TextMark = "Sheet1!A1";

            byte[] tmp = link.Serialize();
            byte[] ser = new byte[tmp.Length - 4];
            Array.Copy(tmp, 4, ser, 0, ser.Length);
            //Assert.AreEqual(data4.Length, ser.Length);
            Assert.IsTrue(Arrays.Equals(data4, ser));
        }
        [Test]
        public void TestCreateEmailtRecord()
        {
            HyperlinkRecord link = new HyperlinkRecord();
            link.CreateUrlLink();
            link.FirstRow = 1;
            link.LastRow = 1;
            link.Label = "email";
            link.Address = "mailto:ebgans@mail.ru?subject=Hello,%20Ebgans!";

            byte[] tmp = link.Serialize();
            byte[] ser = new byte[tmp.Length - 4];
            Array.Copy(tmp, 4, ser, 0, ser.Length);
            Assert.AreEqual(data3.Length, ser.Length);
            Assert.IsTrue(Arrays.Equals(data3, ser));
        }
        [Test]
        public void TestClone()
        {
            byte[][] data = { data1, data2, data3, data4 };
            for (int i = 0; i < data.Length; i++)
            {
                RecordInputStream is1 = TestcaseRecordInputStream.Create(HyperlinkRecord.sid, data[i]);
                HyperlinkRecord link = new HyperlinkRecord(is1);
                HyperlinkRecord clone = (HyperlinkRecord)link.Clone();
                Assert.IsTrue(Arrays.Equals(link.Serialize(), clone.Serialize()));
            }

        }

        [Test]
        public void TestReserializeTargetFrame()
        {
            RecordInputStream in1 = TestcaseRecordInputStream.Create(HyperlinkRecord.sid, dataTargetFrame);
            HyperlinkRecord hr = new HyperlinkRecord(in1);
            byte[] ser = hr.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(HyperlinkRecord.sid, dataTargetFrame, ser);
        }

        [Test]
        public void TestReserializeLinkToWorkbook()
        {

            RecordInputStream in1 = TestcaseRecordInputStream.Create(HyperlinkRecord.sid, dataLinkToWorkbook);
            HyperlinkRecord hr = new HyperlinkRecord(in1);
            byte[] ser = hr.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(HyperlinkRecord.sid, dataLinkToWorkbook, ser);
            if ("YEARFR~1.XLS".Equals(hr.Address))
            {
                throw new AssertionException("Identified bug in reading workbook link");
            }
            Assert.AreEqual("yearfracExamples.xls", hr.Address);
        }
        [Test]
        public void TestReserializeUNC()
        {

            RecordInputStream in1 = TestcaseRecordInputStream.Create(HyperlinkRecord.sid, dataUNC);
            HyperlinkRecord hr = new HyperlinkRecord(in1);
            byte[] ser = hr.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(HyperlinkRecord.sid, dataUNC, ser);
            try
            {
                hr.ToString();
            }
            catch (NullReferenceException)
            {
                throw new AssertionException("Identified bug with option URL and UNC set at same time");
            }
        }
        [Test]
        public void TestGUID()
        {
            GUID g;
            g = GUID.Parse("3F2504E0-4F89-11D3-9A0C-0305E82C3301");
            ConfirmGUID(g, 0x3F2504E0, 0x4F89, 0x11D3, unchecked((long)0x9A0C0305E82C3301L));
            Assert.AreEqual("3F2504E0-4F89-11D3-9A0C-0305E82C3301", g.FormatAsString());

            g = GUID.Parse("13579BDF-0246-8ACE-0123-456789ABCDEF");
            ConfirmGUID(g, 0x13579BDF, 0x0246, 0x8ACE, unchecked((long)0x0123456789ABCDEFL));
            Assert.AreEqual("13579BDF-0246-8ACE-0123-456789ABCDEF", g.FormatAsString());

            byte[] buf = new byte[16];
            g.Serialize(new LittleEndianByteArrayOutputStream(buf, 0));
            String expectedDump = "[DF, 9B, 57, 13, 46, 02, CE, 8A, 01, 23, 45, 67, 89, AB, CD, EF]";
            Assert.AreEqual(expectedDump, HexDump.ToHex(buf));

            // STD Moniker
            g = CreateFromStreamDump("[D0, C9, EA, 79, F9, BA, CE, 11, 8C, 82, 00, AA, 00, 4B, A9, 0B]");
            Assert.AreEqual("79EAC9D0-BAF9-11CE-8C82-00AA004BA90B", g.FormatAsString());
            // URL Moniker
            g = CreateFromStreamDump("[E0, C9, EA, 79, F9, BA, CE, 11, 8C, 82, 00, AA, 00, 4B, A9, 0B]");
            Assert.AreEqual("79EAC9E0-BAF9-11CE-8C82-00AA004BA90B", g.FormatAsString());
            // File Moniker
            g = CreateFromStreamDump("[03, 03, 00, 00, 00, 00, 00, 00, C0, 00, 00, 00, 00, 00, 00, 46]");
            Assert.AreEqual("00000303-0000-0000-C000-000000000046", g.FormatAsString());
        }

        private static GUID CreateFromStreamDump(String s)
        {
            return new GUID(new LittleEndianByteArrayInputStream(HexRead.ReadFromString(s)));
        }

        private void ConfirmGUID(GUID g, int d1, int d2, int d3, long d4)
        {
            Assert.AreEqual(new String(HexDump.IntToHex(d1)), new String(HexDump.IntToHex(g.D1)));
            Assert.AreEqual(new String(HexDump.ShortToHex(d2)), new String(HexDump.ShortToHex(g.D2)));
            Assert.AreEqual(new String(HexDump.ShortToHex(d3)), new String(HexDump.ShortToHex(g.D3)));
            Assert.AreEqual(new String(HexDump.LongToHex(d4)), new String(HexDump.LongToHex(g.D4)));
        }
        [Test]
        public void Test47498()
        {
            RecordInputStream is1 = TestcaseRecordInputStream.Create(HyperlinkRecord.sid, data_47498);
            HyperlinkRecord link = new HyperlinkRecord(is1);
            Assert.AreEqual(2, link.FirstRow);
            Assert.AreEqual(2, link.LastRow);
            Assert.AreEqual(0, link.FirstColumn);
            Assert.AreEqual(0, link.LastColumn);
            ConfirmGUID(HyperlinkRecord.STD_MONIKER, link.Guid);
            ConfirmGUID(HyperlinkRecord.URL_MONIKER, link.Moniker);
            Assert.AreEqual(2, link.LabelOptions);
            int opts = HyperlinkRecord.HLINK_URL | HyperlinkRecord.HLINK_LABEL;
            Assert.AreEqual(opts, link.LinkOptions);
            Assert.AreEqual(0, link.FileOptions);

            Assert.AreEqual("PDF", link.Label);
            Assert.AreEqual("testfolder/test.PDF", link.Address);

            byte[] ser = link.Serialize();
            TestcaseRecordInputStream.ConfirmRecordEncoding(HyperlinkRecord.sid, data_47498, ser);
        }
    }
}