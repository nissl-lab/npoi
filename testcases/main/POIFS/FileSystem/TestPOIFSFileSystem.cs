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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

namespace TestCases.POIFS.FileSystem
{
    using System;
    using System.Collections;
    using System.IO;

    using NUnit.Framework;

    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using NPOI.POIFS.Storage;
    using NPOI.POIFS.Properties;
    using TestCases.HSSF;
    using NPOI.POIFS.Common;
    using System.Collections.Generic;

    /**
     * Tests for POIFSFileSystem
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestPOIFSFileSystem
    {
        private POIDataSamples _samples = POIDataSamples.GetPOIFSInstance();

        /**
         * Mock exception used to ensure correct error handling
         */
        private class MyEx : Exception
        {
            public MyEx()
            {
                // no fields to initialise
            }
        }
        /**
         * Helps facilitate Testing. Keeps track of whether close() was called.
         * Also can throw an exception at a specific point in the stream. 
         */
        private class TestIS : Stream
        {

            private Stream _is;
            private int _FailIndex;
            private int _currentIx;
            private bool _isClosed;

            public TestIS(Stream is1, int FailIndex)
            {
                _is = is1;
                _FailIndex = FailIndex;
                _currentIx = 0;
                _isClosed = false;
            }

            public int Read()
            {
                int result = _is.ReadByte();
                if (result >= 0)
                {
                    CheckRead(1);
                }
                return result;
            }
            public override int Read(byte[] b, int off, int len)
            {
                int result = _is.Read(b, off, len);
                CheckRead(result);
                return result;
            }

            private void CheckRead(int nBytes)
            {
                _currentIx += nBytes;
                if (_FailIndex > 0 && _currentIx > _FailIndex)
                {
                    throw new MyEx();
                }
            }
            public override void Close()
            {
                _isClosed = true;
                _is.Close();
            }
            public bool IsClosed()
            {
                return _isClosed;
            }
            public override void Flush()
            {
                _is.Flush();
            }
            public override long Seek(long offset, SeekOrigin origin)
            {
                return _is.Seek(offset, origin);
            }

            public override void SetLength(long value)
            {
                _is.SetLength(value);
            }
            // Properties
            public override bool CanRead
            {
                get
                {
                    return _is.CanRead;
                }
            }

            public override bool CanSeek
            {
                get
                {
                    return _is.CanSeek;
                }
            }

            public override bool CanWrite
            {
                get
                {
                    return _is.CanWrite;
                }
            }

            public override long Length
            {
                get
                {
                    return _is.Length;
                }
            }

            public override long Position
            {
                get
                {
                    return _is.Position;
                }
                set
                {
                    _is.Position = value;
                }
            }
            public override void Write(byte[] buffer, int offset, int count)
            {
            }
        }

        /**
         * Test for undesired behaviour observable as of svn revision 618865 (5-Feb-2008).
         * POIFSFileSystem was not closing the input stream.
         */
        [Test]
        public void TestAlwaysClose()
        {

            TestIS testIS;

            // Normal case - Read until EOF and close
            testIS = new TestIS(OpenSampleStream("13224.xls"), -1);
            try
            {
                new POIFSFileSystem(testIS);
            }
            catch (IOException)
            {
                throw;
            }
            Assert.IsTrue(testIS.IsClosed(), "input stream was not closed");

            // intended to crash after Reading 10000 bytes
            testIS = new TestIS(OpenSampleStream("13224.xls"), 10000);
            try
            {
                new POIFSFileSystem(testIS);
                Assert.Fail("ex Should have been thrown");
            }
            catch (IOException)
            {
                throw;
            }
            catch (MyEx)
            {
                // expected
            }
            Assert.IsTrue(testIS.IsClosed(), "input stream was not closed"); // but still Should close

        }

        /**
         * Test for bug # 48898 - problem opening an OLE2
         *  file where the last block is short (i.e. not a full
         *  multiple of 512 bytes)
         *  
         * As yet, this problem remains. One school of thought is
         *  not not issue an EOF when we discover the last block
         *  is short, but this seems a bit wrong.
         * The other is to fix the handling of the last block in
         *  POIFS, since it seems to be slight wrong
         */
        [Test]
        public void TestShortLastBlock()
        {
            String[] files = new String[] { "ShortLastBlock.qwp", "ShortLastBlock.wps" };

            for (int i = 0; i < files.Length; i++)
            {
                // Open the file up
                POIFSFileSystem fs = new POIFSFileSystem(
                        _samples.OpenResourceAsStream(files[i])
                );

                // Write it into a temp output array
                MemoryStream baos = new MemoryStream();
                fs.WriteFileSystem(baos);

                // Check sizes
            }
        }

        [Test]
        public void TestFATandDIFATsectors()
        {
            // Open the file up
            try
            {
                Stream stream = _samples.OpenResourceAsStream("ReferencesInvalidSectors.mpp");
                try
                {
                    POIFSFileSystem fs = new POIFSFileSystem(stream);
                    Assert.Fail("File is corrupt and shouldn't have been opened");
                }
                finally
                {
                    stream.Close();
                }
            }
            catch (IOException e)
            {
                String msg = e.Message;
                Assert.IsTrue(msg.StartsWith("Your file contains 695 sectors"));
            }
        }

        [Test]
        public void TestBATandXBAT()
        {
            byte[] hugeStream = new byte[8 * 1024 * 1024];
            POIFSFileSystem fs = new POIFSFileSystem();
            fs.Root.CreateDocument("BIG", new MemoryStream(hugeStream));

            MemoryStream baos = new MemoryStream();
            fs.WriteFileSystem(baos);
            byte[] fsData = baos.ToArray();


            // Check the header was written properly
            Stream inp = new MemoryStream(fsData);
            HeaderBlock header = new HeaderBlock(inp);
            Assert.AreEqual(109 + 21, header.BATCount);
            Assert.AreEqual(1, header.XBATCount);

            ByteBuffer xbatData = ByteBuffer.CreateBuffer(512);
            xbatData.Write(fsData, (1 + header.XBATIndex) * 512, 512);

            xbatData.Position = 0;

            BATBlock xbat = BATBlock.CreateBATBlock(POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS, xbatData);

            for (int i = 0; i < 21; i++)
            {
                Assert.IsTrue(xbat.GetValueAt(i) != POIFSConstants.UNUSED_BLOCK);
            }

            for (int i = 21; i < 127; i++)
                Assert.AreEqual(POIFSConstants.UNUSED_BLOCK, xbat.GetValueAt(i));

            Assert.AreEqual(POIFSConstants.END_OF_CHAIN, xbat.GetValueAt(127));

            RawDataBlockList blockList = new RawDataBlockList(inp, POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS);
            Assert.AreEqual(fsData.Length / 512, blockList.BlockCount() + 1);
            new BlockAllocationTableReader(header.BigBlockSize,
                                            header.BATCount,
                                            header.BATArray,
                                            header.XBATCount,
                                            header.XBATIndex,
                                            blockList);
            Assert.AreEqual(fsData.Length / 512, blockList.BlockCount() + 1);

            //fs = null;
            //fs = new POIFSFileSystem(new MemoryStream(fsData));


            //DirectoryNode root = fs.Root;
            //Assert.AreEqual(1, root.EntryCount);
            //DocumentNode big = (DocumentNode)root.GetEntry("BIG");
            //Assert.AreEqual(hugeStream.Length, big.Size);

        }

        [Test]
        public void Test4KBlocks()
        {

            Stream inp = _samples.OpenResourceAsStream("BlockSize4096.zvi");
            try
            {
                // First up, check that we can process the header properly
                HeaderBlock header_block = new HeaderBlock(inp);
                POIFSBigBlockSize bigBlockSize = header_block.BigBlockSize;
                Assert.AreEqual(4096, bigBlockSize.GetBigBlockSize());

                // Check the fat info looks sane
                Assert.AreEqual(1, header_block.BATArray.Length);
                Assert.AreEqual(1, header_block.BATCount);
                Assert.AreEqual(0, header_block.XBATCount);

                // Now check we can get the basic fat
                RawDataBlockList data_blocks = new RawDataBlockList(inp, bigBlockSize);
                Assert.AreEqual(15, data_blocks.BlockCount());

                // Now try and open properly
                POIFSFileSystem fs = new POIFSFileSystem(
                      _samples.OpenResourceAsStream("BlockSize4096.zvi")
                );
                Assert.IsTrue(fs.Root.EntryCount > 3);

                // Check we can get at all the contents
                CheckAllDirectoryContents(fs.Root);


                // Finally, check we can do a similar 512byte one too
                fs = new POIFSFileSystem(
                     _samples.OpenResourceAsStream("BlockSize512.zvi")
               );
                Assert.IsTrue(fs.Root.EntryCount > 3);
                CheckAllDirectoryContents(fs.Root);
            }
            finally
            {
                inp.Close();
            }
        }


        private void CheckAllDirectoryContents(DirectoryEntry dir)
        {
            IEnumerator<Entry> it = dir.Entries;
            //foreach (Entry entry in dir)
            while (it.MoveNext())
            {
                Entry entry = it.Current;
                if (entry is DirectoryEntry)
                {
                    CheckAllDirectoryContents((DirectoryEntry)entry);
                }
                else
                {
                    DocumentNode doc = (DocumentNode)entry;
                    DocumentInputStream dis = new DocumentInputStream(doc);
                    try
                    {
                        int numBytes = dis.Available();
                        byte[] data = new byte[numBytes];
                        dis.Read(data);
                    }
                    finally
                    {
                        dis.Close();
                    }
                }
            }
        }


        private static Stream OpenSampleStream(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleFileStream(sampleFileName);
        }
    }
}