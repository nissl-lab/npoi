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
using System.IO;
using NPOI.HSSF.Record;
using NPOI.Util;
using NUnit.Framework;

namespace TestCases.HSSF.Record
{
    /**
 * Unit tests for DConRefRecord class.
 *
 * @author Niklas Rehfeld
 */
    public class TestDConRefRecord
    {
        /**
         * record of a proper single-byte external 'volume'-style path with multiple parts and a sheet
         * name.
         */
        byte[] volumeString = new byte[]
    {
        0, 0, 0, 0, 0, 0, //ref (6 bytes) not used...
        17, 0,//cchFile (2 bytes)
        0, //char type
        1, 1, (byte)'c', (byte)'[', (byte)'f', (byte)'o', (byte)'o', 0x3,
        (byte)'b', (byte)'a', (byte)'r', (byte)']', (byte)'s', (byte)'h', (byte)'e', (byte)'e',
        (byte)'t'
    };
        /**
         * record of a proper single-byte external 'unc-volume'-style path with multiple parts and a
         * sheet name.
         */
        byte[] uncVolumeString = new byte[]
    {
        0, 0, 0, 0, 0, 0, //ref (6 bytes) not used...
        34, 0,//cchFile (2 bytes)
        0, //char type
        1, 1, (byte)'@', (byte)'[',(byte) 'c', (byte)'o', (byte)'m', (byte)'p',
        0x3, (byte)'s', (byte)'h', (byte)'a', (byte)'r',(byte) 'e', (byte)'d', 0x3,
        (byte)'r', (byte)'e', (byte)'l', (byte)'a', (byte)'t', (byte)'i', (byte)'v', (byte)'e',
        0x3, (byte)'f', (byte)'o', (byte)'o',(byte) ']', (byte)'s',(byte) 'h',(byte) 'e',
        (byte)'e', (byte)'t'
    };
        /**
         * record of a proper single-byte external 'simple-file-path-dcon' style path with a sheet name.
         */
        byte[] simpleFilePathDconString = new byte[]
    {
        0, 0, 0, 0, 0, 0, //ref (6 bytes) not used...
        16, 0,//cchFile (2 bytes)
        0, //char type
        1, (byte)'c', (byte)'[', (byte)'f', (byte)'o', (byte)'o', 0x3, (byte)'b',
        (byte)'a', (byte)'r', (byte)']', (byte)'s', (byte)'h', (byte)'e', (byte)'e', (byte)'t'
    };
        /**
         * record of a proper 'transfer-protocol'-style path. This one has a sheet name at the end, and
         * another one inside the file path. The spec doesn't seem to care about what they are.
         */
        byte[] transferProtocolString = new byte[]
    {
        0, 0, 0, 0, 0, 0, //ref (6 bytes) not used...
        33, 0,//cchFile (2 bytes)
        0, //char type
        0x1, 0x5, 30, //count = 30
        (byte)'[', (byte)'h', (byte)'t', (byte)'t', (byte)'p', (byte)':',(byte) '/',(byte) '/',
        (byte)'[', (byte)'f', (byte)'o', (byte)'o', 0x3, (byte)'b', (byte)'a', (byte)'r',
        (byte)']', (byte)'s', (byte)'h', (byte)'e', (byte)'e', (byte)'t', (byte)'1', (byte)']',
        (byte)'s', (byte)'h', (byte)'e',(byte) 'e', (byte)'t', (byte)'x'
    };
        /**
         * startup-type path.
         */
        byte[] relVolumeString = new byte[]
    {
        0, 0, 0, 0, 0, 0, //ref (6 bytes) not used...
        16, 0,//cchFile (2 bytes)
        0, //char type
        0x1, 0x2, (byte)'[', (byte)'f', (byte)'o', (byte)'o', 0x3, (byte)'b',
        (byte)'a', (byte)'r', (byte)']', (byte)'s', (byte)'h', (byte)'e', (byte)'e', (byte)'t'
    };
        /**
         * startup-type path.
         */
        byte[] startupString = new byte[]
    {
        0, 0, 0, 0, 0, 0, //ref (6 bytes) not used...
        16, 0,//cchFile (2 bytes)
        0, //char type
        0x1, 0x6, (byte)'[', (byte)'f', (byte)'o', (byte)'o', 0x3, (byte)'b',
        (byte)'a', (byte)'r', (byte)']', (byte)'s', (byte)'h', (byte)'e', (byte)'e', (byte)'t'
    };
        /**
         * alt-startup-type path.
         */
        byte[] altStartupString = new byte[]
    {
        0, 0, 0, 0, 0, 0, //ref (6 bytes) not used...
        16, 0,//cchFile (2 bytes)
        0, //char type
        0x1, 0x7, (byte)'[', (byte)'f', (byte)'o', (byte)'o', 0x3, (byte)'b',
        (byte)'a',(byte) 'r', (byte)']', (byte)'s', (byte)'h', (byte)'e', (byte)'e',(byte) 't'
    };
        /**
         * library-style path.
         */
        byte[] libraryString = new byte[]
    {
        0, 0, 0, 0, 0, 0, //ref (6 bytes) not used...
        16, 0,//cchFile (2 bytes)
        0, //char type
        0x1, 0x8, (byte)'[', (byte)'f', (byte)'o', (byte)'o', 0x3, (byte)'b',
        (byte)'a', (byte)'r', (byte)']', (byte)'s', (byte)'h', (byte)'e', (byte)'e', (byte)'t'
    };
        /**
         * record of single-byte string, external, volume path.
         */
        byte[] data1 = new byte[]
    {
        0, 0, 0, 0, 0, 0, //ref (6 bytes) not used...
        10, 0,//cchFile (2 bytes)
        0, //char type
        1, 1, (byte) 'b', (byte) 'l', (byte) 'a', (byte) ' ', (byte) 't',
        (byte) 'e', (byte) 's', (byte) 't'
    //unused doesn't exist as stFile[1] != 2
    };
        /**
         * record of double-byte string, self-reference.
         */
        byte[] data2 = new byte[]
    {
        0, 0, 0, 0, 0, 0, //ref (6 bytes) not used...
        9, 0,//cchFile (2 bytes)
        1, //char type = unicode
        2, 0, (byte) 'b', 0, (byte) 'l', 0, (byte) 'a', 0, (byte) ' ', 0, (byte) 't', 0,
        (byte) 'e', 0, (byte) 's', (byte) 't', 0,//stFile
        0, 0 //unused (2 bytes as we're using double-byte chars)
    };
        /**
         * record of single-byte string, self-reference.
         */
        byte[] data3 = new byte[]
    {
        0, 0, 0, 0, 0, 0, //ref (6 bytes) not used...
        9, 0,//cchFile (2 bytes)
        0, //char type = ansi
        2, (byte) 'b', (byte) 'l', (byte) 'a', (byte) ' ', (byte) 't', (byte) 'e', (byte) 's',
        (byte) 't',//stFile
        0 //unused (1 byte as we're using single byes)
    };
        /**
         * double-byte string, external reference, unc-volume.
         */
        byte[] data4 = new byte[]
    {
        0, 0, 0, 0, 0, 0, //ref (6 bytes) not used...
        16, 0,//cchFile (2 bytes)
        //stFile starts here:
        1, //char type = unicode
        1, 0, 1, 0, 0x40, 0, (byte) 'c', 0, (byte) 'o', 0, (byte) 'm', 0, (byte) 'p', 0, 0x03, 0,
        (byte) 'b', 0, (byte) 'l', 0, (byte) 'a', 0, 0x03, 0, (byte) 't', 0, (byte) 'e', 0,
        (byte) 's', 0, (byte) 't', 0,
    //unused doesn't exist as stFile[1] != 2
    };

        /**
         * test read-constructor-then-serialize for a single-byte external reference strings of
         * various flavours. This uses the RecordInputStream constructor.
         * @throws IOException
         */
        [Test]
        public void TestReadWriteSbExtRef()
        {
            TestReadWrite(data1, "read-write single-byte external reference, volume type path");
            TestReadWrite(volumeString,
                    "read-write properly formed single-byte external reference, volume type path");
            TestReadWrite(uncVolumeString,
                    "read-write properly formed single-byte external reference, UNC volume type path");
            TestReadWrite(relVolumeString,
                    "read-write properly formed single-byte external reference, rel-volume type path");
            TestReadWrite(simpleFilePathDconString,
                    "read-write properly formed single-byte external reference, simple-file-path-dcon type path");
            TestReadWrite(transferProtocolString,
                    "read-write properly formed single-byte external reference, transfer-protocol type path");
            TestReadWrite(startupString,
                    "read-write properly formed single-byte external reference, startup type path");
            TestReadWrite(altStartupString,
                    "read-write properly formed single-byte external reference, alt-startup type path");
            TestReadWrite(libraryString,
                    "read-write properly formed single-byte external reference, library type path");
        }

        /**
         * test read-constructor-then-serialize for a double-byte external reference 'UNC-Volume' style
         * string
         * <p/>
         * @throws IOException
         */
        [Test]
        public void TestReadWriteDbExtRefUncVol()
        {
            TestReadWrite(data4, "read-write double-byte external reference, UNC volume type path");
        }

        private void TestReadWrite(byte[] data, String message)
        {
            RecordInputStream is1 = TestcaseRecordInputStream.Create(81, data);
            DConRefRecord d = new DConRefRecord(is1);
            MemoryStream bos = new MemoryStream(data.Length);
            LittleEndianOutputStream o = new LittleEndianOutputStream(bos);
            d.Serialize(o);
            o.Flush();

            Assert.IsTrue(Arrays.Equals(data, bos.ToArray()), message);
        }

        /**
         * test read-constructor-then-serialize for a double-byte self-reference style string
         * <p/>
         * @throws IOException
         */
        [Test]
        public void TestReadWriteDbSelfRef()
        {
            TestReadWrite(data2, "read-write double-byte self reference");
        }

        /**
         * test read-constructor-then-serialize for a single-byte self-reference style string
         * <p/>
         * @throws IOException
         */
        [Test]
        public void TestReadWriteSbSelfRef()
        {
            TestReadWrite(data3, "read-write single byte self reference");
        }

        /**
         * Test of getDataSize method, of class DConRefRecord.
         */
        [Test]
        public void TestGetDataSize()
        {
            DConRefRecord instance = new DConRefRecord(TestcaseRecordInputStream.Create(81, data1));
            int expResult = data1.Length;
            int result = instance.RecordSize - 4;
            Assert.AreEqual(expResult, result, "single byte external reference, volume type path data size");
            instance = new DConRefRecord(TestcaseRecordInputStream.Create(81, data2));
            Assert.AreEqual(data2.Length, instance.RecordSize - 4, "double byte self reference data size");
            instance = new DConRefRecord(TestcaseRecordInputStream.Create(81, data3));
            Assert.AreEqual(data3.Length, instance.RecordSize - 4, "single byte self reference data size");
            instance = new DConRefRecord(TestcaseRecordInputStream.Create(81, data4));
            Assert.AreEqual(data4.Length, instance.RecordSize - 4, "double byte external reference, UNC volume type path data size");
        }

        /**
         * Test of getSid method, of class DConRefRecord.
         */
        [Test]
        public void TestGetSid()
        {
            DConRefRecord instance = new DConRefRecord(TestcaseRecordInputStream.Create(81, data1));
            short expResult = 81;
            short result = instance.Sid;
            Assert.AreEqual(expResult, result, "SID");
        }

        /**
         * Test of getPath method, of class DConRefRecord.
         * @todo different types of paths.
         */
        [Test]
        public void TestGetPath()
        {
            DConRefRecord instance = new DConRefRecord(TestcaseRecordInputStream.Create(81, data1));
            byte[] expResult = Arrays.CopyOfRange(data1, 9, data1.Length);
            byte[] result = instance.GetPath();
            Assert.IsTrue(Arrays.Equals(expResult, result), "get path");
        }

        /**
         * Test of isExternalRef method, of class DConRefRecord.
         */
        [Test]
        public void TestIsExternalRef()
        {
            DConRefRecord instance = new DConRefRecord(TestcaseRecordInputStream.Create(81, data1));
            Assert.IsTrue(instance.IsExternalRef, "external reference");
            instance = new DConRefRecord(TestcaseRecordInputStream.Create(81, data2));
            Assert.IsFalse(instance.IsExternalRef, "internal reference");
        }
    }
}