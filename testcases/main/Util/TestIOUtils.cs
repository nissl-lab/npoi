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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.Util
{
    using NPOI;
    using NPOI.Util;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    /// <summary>
    /// Class to test IOUtils
    /// </summary>
    [TestFixture]
    public class TestIOUtils
    {

        static FileInfo TMP = null;
        static long LENGTH = new Random().Next(10000);

        [SetUp]
        public void SetUp()
        {
            TMP = TempFile.CreateTempFile("poi-ioutils-", "");
            Stream os = TMP.Open(FileMode.OpenOrCreate);
            for(int i = 0; i < LENGTH; i++)
            {
                os.WriteByte(0x01);
            }
            os.Flush();
            os.Close();

        }

        [TearDown]
        public void TearDown()
        {
            //noinspection ResultOfMethodCallIgnored
            TMP.Delete();
        }

        [Test]
        public void TestPeekFirst8Bytes()
        {
            ClassicAssert.IsTrue(Arrays.Equals(Encoding.UTF8.GetBytes("01234567"),
                    IOUtils.PeekFirst8Bytes(new ByteArrayInputStream(Encoding.UTF8.GetBytes("01234567")))));
        }

        [Test]
        public void TestPeekFirst8BytesWithPushbackInputStream()
        {
            ClassicAssert.IsTrue(Arrays.Equals(Encoding.UTF8.GetBytes("01234567"),
                    IOUtils.PeekFirst8Bytes(new PushbackInputStream(new ByteArrayInputStream(Encoding.UTF8.GetBytes("01234567")), 8))));
        }

        [Test]
        public void TestPeekFirst8BytesTooLessAvailable()
        {
            ClassicAssert.IsTrue(Arrays.Equals(new byte[] { 1, 2, 3, 0, 0, 0, 0, 0 },
                    IOUtils.PeekFirst8Bytes(new ByteArrayInputStream(new byte[] { 1, 2, 3 })))
            );
        }

        [Test]
        public void TestPeekFirst8BytesEmpty()
        {
            Assert.Throws<EmptyFileException>(() =>
                IOUtils.PeekFirst8Bytes(new ByteArrayInputStream(new byte[] { }))
            );
        }

        [Test]
        public void TestToByteArray()
        {
            ClassicAssert.IsTrue(Arrays.Equals(new byte[] { 1, 2, 3 },
                 IOUtils.ToByteArray(new MemoryStream(new byte[] { 1, 2, 3 })))
            );
        }

        [Test]
        public void TestToByteArrayToSmall()
        {
            ClassicAssert.Throws<IOException>(() =>
            ClassicAssert.IsTrue(Arrays.Equals(new byte[] { 1, 2, 3 },
                    IOUtils.ToByteArray(new MemoryStream(new byte[] { 1, 2, 3 }), 10)))
            );
        }

        [Test]
        public void TestToByteArrayByteBuffer()
        {
            ClassicAssert.IsTrue(Arrays.Equals(new byte[] { 1, 2, 3 },
                    IOUtils.ToByteArray(new ByteBuffer(new byte[] { 1, 2, 3 }, 0, 3), 10)));
        }

        [Test]
        public void TestToByteArrayByteBufferToSmall()
        {
            ClassicAssert.IsTrue(Arrays.Equals(new byte[] { 1, 2, 3, 4, 5, 6, 7 },
                    IOUtils.ToByteArray(new ByteBuffer(new byte[] { 1, 2, 3, 4, 5, 6, 7 }, 0, 7), 3)));
        }

        [Test]
        public void TestSkipFully()
        {
            using InputStream is1 =  new FileInputStream(TMP.Open(FileMode.OpenOrCreate));
            long skipped = IOUtils.SkipFully(is1, 20000L);
            ClassicAssert.AreEqual(LENGTH, skipped, "length: "+LENGTH);
        }

        [Test]
        public void TestSkipFullyGtIntMax()
        {
            using InputStream is1 =  new FileInputStream(TMP.Open(FileMode.OpenOrCreate));
            long skipped = IOUtils.SkipFully(is1, Int32.MaxValue + 20000L);
            ClassicAssert.AreEqual(LENGTH, skipped, "length: "+LENGTH);
        }

        [Test]
        public void TestSkipFullyByteArray()
        {
            using MemoryStream bos = new MemoryStream();
            using InputStream is1 = new FileInputStream(TMP.Open(FileMode.OpenOrCreate));
            IOUtils.Copy(is1, bos);
            long skipped = IOUtils.SkipFully(new ByteArrayInputStream(bos.ToArray()), 20000L);
            ClassicAssert.AreEqual(LENGTH, skipped, "length: "+LENGTH);
        }

        [Test]
        public void TestSkipFullyByteArrayGtIntMax()
        {
            using MemoryStream bos = new MemoryStream();
            using InputStream is1 = new FileInputStream(TMP.Open(FileMode.OpenOrCreate));
            IOUtils.Copy(is1, bos);
            long skipped = IOUtils.SkipFully(new ByteArrayInputStream(bos.ToArray()), Int32.MaxValue+ 20000L);
            ClassicAssert.AreEqual(LENGTH, skipped, "length: "+LENGTH);
        }

        [Test]
        public void TestSkipFullyBug61294()
        {
            IOUtils.SkipFully(new ByteArrayInputStream(new byte[0]), 1);
        }

        [Test]
        public void TestZeroByte()
        {
            long skipped = IOUtils.SkipFully((new ByteArrayInputStream(new byte[0])), 100);
            ClassicAssert.AreEqual(-1L, skipped, "zero byte");
        }

        [Test]
        public void TestSkipZero()
        {
            using InputStream is1 =  new FileInputStream(TMP.Open(FileMode.OpenOrCreate));
            long skipped = IOUtils.SkipFully(is1, 0);
            ClassicAssert.AreEqual(0, skipped, "zero length");
        }
        [Test]
        public void TestSkipNegative()
        {
            ClassicAssert.Throws<ArgumentException>(() =>
            {
                using InputStream is1 =  new FileInputStream(TMP.Open(FileMode.OpenOrCreate));
                IOUtils.SkipFully(is1, -1);
            });
        }

        [Test]
        public void TestWonkyInputStream()
        {
            long skipped = IOUtils.SkipFully(new WonkyInputStream(), 10000);
            ClassicAssert.AreEqual(10000, skipped, "length: "+LENGTH);
        }

        /// <summary>
        /// This returns 0 for the first call to skip and then reads
        /// as requested.  This tests that the fallback to read() works.
        /// </summary>
        private class WonkyInputStream : InputStream
        {
            int skipCalled = 0;
            int readCalled = 0;

            public override bool CanRead => throw new NotImplementedException();

            public override bool CanSeek => throw new NotImplementedException();

            public override long Length => throw new NotImplementedException();

            public override long Position { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

            public override int Read()
            {
                readCalled++;
                return 0;
            }
            public override int Read(byte[] arr, int offset, int len)
            {
                readCalled++;
                return len;
            }
            public override long Skip(long len)
            {
                skipCalled++;
                if(skipCalled == 1)
                {
                    return 0;
                }
                else if(skipCalled > 100)
                {
                    return len;
                }
                else
                {
                    return 100;
                }
            }
            public override int Available()
            {
                return 100000;
            }

            public override void Flush()
            {
                throw new NotImplementedException();
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                throw new NotImplementedException();
            }

            public override void SetLength(long value)
            {
                throw new NotImplementedException();
            }
        }
    }
}
