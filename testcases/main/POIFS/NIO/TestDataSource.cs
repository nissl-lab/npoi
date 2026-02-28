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

using System;
using System.Text;
using System.Collections.Generic;

using NUnit.Framework;using NUnit.Framework.Legacy;
using NPOI.POIFS.NIO;
using System.IO;
using NPOI.Util;

namespace TestCases.POIFS.NIO
{
    /// <summary>
    /// Summary description for TestDataSource
    /// </summary>
    [TestFixture]
    public class TestDataSource
    {
        private static POIDataSamples data = POIDataSamples.GetPOIFSInstance();
        public TestDataSource()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestFile()
        {
            FileStream f = data.GetFile("Notes.ole2");

            FileBackedDataSource ds = new FileBackedDataSource(f, false);

            try
            {
                CheckDataSource(ds, false);
            }
            finally
            {
                ds.Close();
            }

            // try a second time
            ds = new FileBackedDataSource(f, false);
            try
            {
                CheckDataSource(ds, false);
            }
            finally
            {
                ds.Close();
            }
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestFileWritable()
        {
            FileInfo temp = TempFile.CreateTempFile("TestDataSource", ".test");
            try
            {
                WriteDataToFile(temp);

                FileBackedDataSource ds = new FileBackedDataSource(temp, false);
                try
                {
                    CheckDataSource(ds, true);
                }
                finally
                {
                    ds.Close();
                }

                // try a second time
                ds = new FileBackedDataSource(temp, false);
                try
                {
                    CheckDataSource(ds, true);
                }
                finally
                {
                    ds.Close();
                }

                WriteDataToFile(temp);
            }
            finally
            {
                ClassicAssert.IsTrue(temp.Exists);
                temp.Delete();
                ClassicAssert.IsTrue(!File.Exists(temp.FullName), "Could not delete file " + temp);
            }
        }

        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestRewritableFile()
        {
            FileInfo temp = TempFile.CreateTempFile("TestDataSource", ".test");
            try
            {
                WriteDataToFile(temp);

                FileBackedDataSource ds = new FileBackedDataSource(temp, true);
                try
                {
                    ByteBuffer buf = ds.Read(0, 10);
                    ClassicAssert.IsNotNull(buf);
                    buf = ds.Read(8, 0x400);
                    ClassicAssert.IsNotNull(buf);
                }
                finally
                {
                    ds.Close();
                }

                // try a second time
                ds = new FileBackedDataSource(temp, true);
                try
                {
                    ByteBuffer buf = ds.Read(0, 10);
                    ClassicAssert.IsNotNull(buf);
                    buf = ds.Read(8, 0x400);
                    ClassicAssert.IsNotNull(buf);
                }
                finally
                {
                    ds.Close();
                }

                WriteDataToFile(temp);
            }
            finally
            {
                ClassicAssert.IsTrue(temp.Exists);
                temp.Delete();
                ClassicAssert.IsTrue(!File.Exists(temp.FullName));
            }
        }

        private void WriteDataToFile(FileInfo temp)
        {
            FileStream str = temp.Create();
            try
            {
                Stream in1 = data.OpenResourceAsStream("Notes.ole2");
                try
                {
                    IOUtils.Copy(in1, str);
                }
                finally
                {
                    in1.Close();
                }
            }
            finally
            {
                str.Close();
            }
        }

        private void CheckDataSource(FileBackedDataSource ds, bool writeable)
        {
            ClassicAssert.AreEqual(writeable, ds.IsWriteable);
            //ClassicAssert.IsNotNull(ds.Channel);

            // rewriting changes the size
            if (writeable)
            {
                ClassicAssert.IsTrue(ds.Size == 8192 || ds.Size == 8198, "Had: " + ds.Size);
            }
            else
            {
                ClassicAssert.AreEqual(8192, ds.Size);
            }


            ByteBuffer bs;
            bs = ds.Read(4, 0);
            ClassicAssert.AreEqual(4, bs.Length);
            //ClassicAssert.AreEqual(0, bs
            ClassicAssert.AreEqual(unchecked((byte)0xd0 - (byte)256), bs[0]);
            ClassicAssert.AreEqual(unchecked((byte)0xcf - (byte)256), bs[1]);
            ClassicAssert.AreEqual(unchecked((byte)0x11 - 0), bs[2]);
            ClassicAssert.AreEqual(unchecked((byte)0xe0 - (byte)256), bs[3]);

            bs = ds.Read(8, 0x400);
            ClassicAssert.AreEqual(8, bs.Length);
            //ClassicAssert.AreEqual(0, bs.position());
            ClassicAssert.AreEqual((byte)'R', bs[0]);
            ClassicAssert.AreEqual(0, bs[1]);
            ClassicAssert.AreEqual((byte)'o', bs[2]);
            ClassicAssert.AreEqual(0, bs[3]);
            ClassicAssert.AreEqual((byte)'o', bs[4]);
            ClassicAssert.AreEqual(0, bs[5]);
            ClassicAssert.AreEqual((byte)'t', bs[6]);
            ClassicAssert.AreEqual(0, bs[7]);

            // Can go to the end, but not past it
            bs = ds.Read(8, 8190);
            ClassicAssert.AreEqual(0, bs.Position);// TODO How best to warn of a short read?

            // Can't go off the end
            try
            {
                bs = ds.Read(4, 8192);
                if (!writeable)
                {
                    Assert.Fail("Shouldn't be able to read off the end of the file");
                }
            }
            catch (IndexOutOfRangeException)
            {
            }
        }


        [Test]
        public void TestByteArray()
        {
            byte[] data = new byte[256];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)i;
            }

            ByteArrayBackedDataSource ds = new ByteArrayBackedDataSource(data);

            ByteBuffer bs;
            bs = ds.Read(4, 0);
            //assertEquals(0, bs.position());
            ClassicAssert.AreEqual(0x00, bs[0]);
            ClassicAssert.AreEqual(0x01, bs[1]);
            ClassicAssert.AreEqual(0x02, bs[2]);
            ClassicAssert.AreEqual(0x03, bs[3]);


            bs = ds.Read(4, 100);
            ClassicAssert.AreEqual(100, bs.Position);
            ClassicAssert.AreEqual(100, bs.Read());
            ClassicAssert.AreEqual(101, bs.Read());
            ClassicAssert.AreEqual(102, bs.Read());
            ClassicAssert.AreEqual(103, bs.Read());

            bs = ds.Read(4, 252);
            ClassicAssert.AreEqual(unchecked((byte)-4), bs.Read());
            ClassicAssert.AreEqual(unchecked((byte)-3), bs.Read());
            ClassicAssert.AreEqual(unchecked((byte)-2), bs.Read());
            ClassicAssert.AreEqual(unchecked((byte)-1), bs.Read());

            // Off the end
            bs = ds.Read(4, 254);
            ClassicAssert.AreEqual(unchecked((byte)-2), bs.Read());
            ClassicAssert.AreEqual(unchecked((byte)-1), bs.Read());
            try
            {
                //bs.get();
                //fail("Shouldn't be able to read off the end");
            }
            catch (Exception) { }

            // Past the end
            try
            {
                ds.Read(4, 256);
                Assert.Fail("Shouldn't be able to read off the end");
            }
            catch (IndexOutOfRangeException) { }


            // Overwrite
            bs = ByteBuffer.CreateBuffer(4);
            //bs.put(0, (byte)-55);
            //bs.put(1, (byte)-54);
            //bs.put(2, (byte)-53);
            //bs.put(3, (byte)-52);
            //bs.Write(0, unchecked((byte)-55));
            //bs.Write(1, unchecked((byte)-54));
            //bs.Write(2, unchecked((byte)-53));
            //bs.Write(3, unchecked((byte)-52));
            bs[0] = unchecked((byte)-55);
            bs[1] = unchecked((byte)-54);
            bs[2] = unchecked((byte)-53);
            bs[3] = unchecked((byte)-52);

            ClassicAssert.AreEqual(256, ds.Size);
            ds.Write(bs, 40);
            ClassicAssert.AreEqual(256, ds.Size);
            bs = ds.Read(4, 40);

            ClassicAssert.AreEqual(unchecked((byte)-55), bs.Read());
            ClassicAssert.AreEqual(unchecked((byte)-54), bs.Read());
            ClassicAssert.AreEqual(unchecked((byte)-53), bs.Read());
            ClassicAssert.AreEqual(unchecked((byte)-52), bs.Read());

            // Append
            //bs = ByteBuffer.allocate(4);
            bs = ByteBuffer.CreateBuffer(4);
            //bs.Write(0, unchecked((byte)-55));
            //bs.Write(1, unchecked((byte)-54));
            //bs.Write(2, unchecked((byte)-53));
            //bs.Write(3, unchecked((byte)-52));
            bs[0] = unchecked((byte)-55);
            bs[1] = unchecked((byte)-54);
            bs[2] = unchecked((byte)-53);
            bs[3] = unchecked((byte)-52);

            //bs[0] = unchecked((byte)-55);
            //bs[1] = unchecked((byte)-54);
            //bs[2] = unchecked((byte)-53);
            //bs[3] = unchecked((byte)-52);

            ClassicAssert.AreEqual(256, ds.Size);
            ds.Write(bs, 256);
            ClassicAssert.AreEqual(260, ds.Size);

            bs = ds.Read(4, 256);
            ClassicAssert.AreEqual(256, bs.Position);
            ClassicAssert.AreEqual(unchecked((byte)-55), bs.Read());
            ClassicAssert.AreEqual(unchecked((byte)-54), bs.Read());
            ClassicAssert.AreEqual(unchecked((byte)-53), bs.Read());
            ClassicAssert.AreEqual(unchecked((byte)-52), bs.Read());

        }
    }
}
