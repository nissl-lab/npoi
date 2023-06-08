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
        

using System;
using System.Collections;
using System.IO;

using NUnit.Framework;

using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NPOI.POIFS.Storage;
using NPOI.POIFS.Properties;
using NPOI.POIFS.NIO;

namespace TestCases.POIFS.FileSystem
{
    /**
     * Class to Test POIFSDocumentReader functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestDocumentInputStream
    {
        //private DocumentNode _workbook;
        private DocumentNode _workbook_n;
        private DocumentNode _workbook_o;
        private byte[] _workbook_data;
        private static int _workbook_size = 5000;
        private static int _buffer_size = 6;

        public TestDocumentInputStream()
        {

            int blocks = (_workbook_size + 511) / 512;

            _workbook_data = new byte[512 * blocks];
            Arrays.Fill(_workbook_data, unchecked((byte)-1));
            for (int j = 0; j < _workbook_size; j++)
            {
                _workbook_data[j] = (byte)(j * j);
            }

            // Create the Old POIFS Version
            RawDataBlock[] rawBlocks = new RawDataBlock[blocks];
            MemoryStream stream =
                new MemoryStream(_workbook_data);

            for (int j = 0; j < blocks; j++)
            {
                rawBlocks[j] = new RawDataBlock(stream);
            }
            OPOIFSDocument document = new OPOIFSDocument("Workbook", rawBlocks,
                                                       _workbook_size);

            _workbook_o = new DocumentNode(
                document.DocumentProperty,
                new DirectoryNode(
                    new DirectoryProperty("Root Entry"), (POIFSFileSystem)null, null));

            // Now create the NPOIFS Version
            byte[] _workbook_data_only = new byte[_workbook_size];
            Array.Copy(_workbook_data, 0, _workbook_data_only, 0, _workbook_size);

            NPOIFSFileSystem npoifs = new NPOIFSFileSystem();
            // Make it easy when debugging to see what isn't the doc
            byte[] minus1 = new byte[512];
            Arrays.Fill(minus1, unchecked((byte)-1));
            npoifs.GetBlockAt(-1).Write(minus1);
            npoifs.GetBlockAt(0).Write(minus1);
            npoifs.GetBlockAt(1).Write(minus1);

            // Create the NPOIFS document
            _workbook_n = (DocumentNode)npoifs.CreateDocument(
                  new MemoryStream(_workbook_data_only),
                  "Workbook"
            );
        }

        /**
         * Test constructor
         *
         * @exception IOException
         */
        [Test]
        public void TestConstructor()
        {
            DocumentInputStream ostream = new ODocumentInputStream(_workbook_o);
            DocumentInputStream nstream = new NDocumentInputStream(_workbook_n);

            Assert.AreEqual(_workbook_size, _workbook_o.Size);
            Assert.AreEqual(_workbook_size, _workbook_n.Size);
            Assert.AreEqual(_workbook_size, ostream.Available());
            Assert.AreEqual(_workbook_size, nstream.Available());

            ostream.Close();
            nstream.Close();
        }

        /**
         * Test Available behavior
         *
         * @exception IOException
         */
        [Test]
        public void TestAvailable()
        {
            DocumentInputStream ostream = new DocumentInputStream(_workbook_o);
            DocumentInputStream nstream = new NDocumentInputStream(_workbook_n);

            Assert.AreEqual(_workbook_size, ostream.Available());
            Assert.AreEqual(_workbook_size, nstream.Available());
            ostream.Close();
            nstream.Close();

            try
            {
                ostream.Available();
                Assert.Fail("Should have caught IOException");
            }
            catch (InvalidOperationException)
            {
                // as expected
            }
            try
            {
                nstream.Available();
                Assert.Fail("Should have caught IOException");
            }
            catch (InvalidOperationException)
            {
                // as expected
            }
        }

        /**
         * Test mark/reset/markSupported.
         *
         * @exception IOException
         */
        [Test]
        public void TestMarkFunctions()
        {
            byte[] buffer = new byte[_workbook_size / 5];
            byte[] small_buffer = new byte[212];

            DocumentInputStream[] streams = new DocumentInputStream[] {
              new DocumentInputStream(_workbook_o),
              new NDocumentInputStream(_workbook_n)
        };
            foreach (DocumentInputStream stream in streams)
            {
                // Read a fifth of it, and check all's correct
                stream.Read(buffer);
                for (int j = 0; j < buffer.Length; j++)
                {
                    Assert.AreEqual(_workbook_data[j], buffer[j], "Checking byte " + j);
                }
                Assert.AreEqual(_workbook_size - buffer.Length, stream.Available());

                // Reset, and check the available goes back to being the
                //  whole of the stream
                stream.Reset();
                Assert.AreEqual(_workbook_size, stream.Available());


                // Read part of a block
                stream.Read(small_buffer);
                for (int j = 0; j < small_buffer.Length; j++)
                {
                    Assert.AreEqual(_workbook_data[j], small_buffer[j], "Checking byte " + j);
                }
                Assert.AreEqual(_workbook_size - small_buffer.Length, stream.Available());
                stream.Mark(0);

                // Read the next part
                stream.Read(small_buffer);
                for (int j = 0; j < small_buffer.Length; j++)
                {
                    Assert.AreEqual(_workbook_data[j + small_buffer.Length], small_buffer[j], "Checking byte " + j);
                }
                Assert.AreEqual(_workbook_size - 2 * small_buffer.Length, stream.Available());

                // Reset, check it goes back to where it was
                stream.Reset();
                Assert.AreEqual(_workbook_size - small_buffer.Length, stream.Available());

                // Read 
                stream.Read(small_buffer);
                for (int j = 0; j < small_buffer.Length; j++)
                {
                    Assert.AreEqual(_workbook_data[j + small_buffer.Length], small_buffer[j], "Checking byte " + j);
                }
                Assert.AreEqual(_workbook_size - 2 * small_buffer.Length, stream.Available());


                // Now read at various points
                Arrays.Fill(small_buffer, (byte)0);
                stream.Read(small_buffer, 6, 8);
                stream.Read(small_buffer, 100, 10);
                stream.Read(small_buffer, 150, 12);
                int pos = small_buffer.Length * 2;
                for (int j = 0; j < small_buffer.Length; j++)
                {
                    byte exp = 0;
                    if (j >= 6 && j < 6 + 8)
                    {
                        exp = _workbook_data[pos];
                        pos++;
                    }
                    if (j >= 100 && j < 100 + 10)
                    {
                        exp = _workbook_data[pos];
                        pos++;
                    }
                    if (j >= 150 && j < 150 + 12)
                    {
                        exp = _workbook_data[pos];
                        pos++;
                    }

                    Assert.AreEqual(exp, small_buffer[j], "Checking byte " + j);
                }
            }

            // Now repeat it with spanning multiple blocks
            streams = new DocumentInputStream[] {
              new DocumentInputStream(_workbook_o),
              new NDocumentInputStream(_workbook_n)        };
            foreach (DocumentInputStream stream in streams)
            {
                // Read several blocks work
                buffer = new byte[_workbook_size / 5];
                stream.Read(buffer);
                for (int j = 0; j < buffer.Length; j++)
                {
                    Assert.AreEqual(_workbook_data[j], buffer[j], "Checking byte " + j);
                }
                Assert.AreEqual(_workbook_size - buffer.Length, stream.Available());

                // Read all of it again, check it began at the start again
                stream.Reset();
                Assert.AreEqual(_workbook_size, stream.Available());

                stream.Read(buffer);
                for (int j = 0; j < buffer.Length; j++)
                {
                    Assert.AreEqual(_workbook_data[j], buffer[j], "Checking byte " + j);
                }

                // Mark our position, and read another whole buffer
                stream.Mark(12);
                stream.Read(buffer);
                Assert.AreEqual(_workbook_size - (2 * buffer.Length),
                      stream.Available());
                for (int j = buffer.Length; j < (2 * buffer.Length); j++)
                {
                    Assert.AreEqual(_workbook_data[j], buffer[j - buffer.Length], "Checking byte " + j);
                }

                // Reset, should go back to only one buffer full read
                stream.Reset();
                Assert.AreEqual(_workbook_size - buffer.Length, stream.Available());

                // Read the buffer again
                stream.Read(buffer);
                Assert.AreEqual(_workbook_size - (2 * buffer.Length),
                      stream.Available());
                for (int j = buffer.Length; j < (2 * buffer.Length); j++)
                {
                    Assert.AreEqual(_workbook_data[j], buffer[j - buffer.Length], "Checking byte " + j);
                }
                Assert.IsTrue(stream.MarkSupported());
            }
        }

        /**
         * Test simple Read method
         *
         * @exception IOException
         */
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestReadSingleByte()
        {
            DocumentInputStream[] streams = new DocumentInputStream[] {
             new DocumentInputStream(_workbook_o),
             new NDocumentInputStream(_workbook_n)
       };
            foreach (DocumentInputStream stream in streams)
            {
                int remaining = _workbook_size;

                // Try and read each byte in turn
                for (int j = 0; j < _workbook_size; j++)
                {
                    int b = stream.Read();
                    Assert.IsTrue(b >= 0, "Checking sign of " + j);
                    Assert.AreEqual(_workbook_data[j],
                          (byte)b, "validating byte " + j);
                    remaining--;
                    Assert.AreEqual(
                          remaining, stream.Available(), "Checking remaining After Reading byte " + j);
                }

                // Ensure we fell off the end
                Assert.AreEqual(-1, stream.Read());

                // Check that After close we can no longer read
                stream.Close();
                try
                {
                    stream.Read();
                    Assert.Fail("Should have caught IOException");
                }
                catch (IOException)
                {
                    // as expected
                }
            }
        }

        /**
         * Test buffered Read
         *
         * @exception IOException
         */
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestBufferRead()
        {
            DocumentInputStream[] streams = new DocumentInputStream[] {
             new DocumentInputStream(_workbook_o),
             new NDocumentInputStream(_workbook_n)
       };
            foreach (DocumentInputStream stream in streams)
            {
                // Need to give a byte array to read
                try
                {
                    stream.Read(null);
                    Assert.Fail("Should have caught NullPointerException");
                }
                catch (NullReferenceException)
                {
                    // as expected
                }

                // test Reading zero length buffer
                Assert.AreEqual(0, stream.Read(new byte[0]));
                Assert.AreEqual(_workbook_size, stream.Available());
                byte[] buffer = new byte[_buffer_size];
                int offset = 0;

                while (stream.Available() >= buffer.Length)
                {
                    Assert.AreEqual(_buffer_size, stream.Read(buffer));
                    for (int j = 0; j < buffer.Length; j++)
                    {
                        Assert.AreEqual(
                              _workbook_data[offset], buffer[j], "in main loop, byte " + offset);
                        offset++;
                    }
                    Assert.AreEqual(_workbook_size - offset,
                          stream.Available(), "offset " + offset);
                }
                Assert.AreEqual(_workbook_size % _buffer_size, stream.Available());
                Arrays.Fill(buffer, (byte)0);
                int count = stream.Read(buffer);

                Assert.AreEqual(_workbook_size % _buffer_size, count);
                for (int j = 0; j < count; j++)
                {
                    Assert.AreEqual(
                          _workbook_data[offset], buffer[j], "past main loop, byte " + offset);
                    offset++;
                }
                Assert.AreEqual(_workbook_size, offset);
                for (int j = count; j < buffer.Length; j++)
                {
                    Assert.AreEqual(0, buffer[j], "Checking remainder, byte " + j);
                }
                Assert.AreEqual(-1, stream.Read(buffer));
                stream.Close();
                try
                {
                    stream.Read(buffer);
                    Assert.Fail("Should have caught IOException");
                }
                catch (IOException)
                {
                    // as expected
                }
            }
        }

        /**
         * Test complex buffered Read
         *
         * @exception IOException
         */
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestComplexBufferRead()
        {
            DocumentInputStream[] streams = new DocumentInputStream[] {
             new DocumentInputStream(_workbook_o),
             new NDocumentInputStream(_workbook_n)
       };
            foreach (DocumentInputStream stream in streams)
            {
                try
                {
                    stream.Read(null, 0, 1);
                    Assert.Fail("Should have caught NullPointerException");
                }
                catch (ArgumentException)
                {
                    // as expected
                }

                // test illegal offsets and lengths
                try
                {
                    stream.Read(new byte[5], -4, 0);
                    Assert.Fail("Should have caught IndexOutOfBoundsException");
                }
                catch (IndexOutOfRangeException)
                {
                    // as expected
                }
                try
                {
                    stream.Read(new byte[5], 0, -4);
                    Assert.Fail("Should have caught IndexOutOfBoundsException");
                }
                catch (IndexOutOfRangeException)
                {
                    // as expected
                }
                try
                {
                    stream.Read(new byte[5], 0, 6);
                    Assert.Fail("Should have caught IndexOutOfBoundsException");
                }
                catch (IndexOutOfRangeException)
                {
                    // as expected
                }

                // test Reading zero
                Assert.AreEqual(0, stream.Read(new byte[5], 0, 0));
                Assert.AreEqual(_workbook_size, stream.Available());
                byte[] buffer = new byte[_workbook_size];
                int offset = 0;

                while (stream.Available() >= _buffer_size)
                {
                    Arrays.Fill(buffer, (byte)0);
                    Assert.AreEqual(_buffer_size,
                          stream.Read(buffer, offset, _buffer_size));
                    for (int j = 0; j < offset; j++)
                    {
                        Assert.AreEqual(0, buffer[j], "Checking byte " + j);
                    }
                    for (int j = offset; j < (offset + _buffer_size); j++)
                    {
                        Assert.AreEqual(_workbook_data[j],
                              buffer[j], "Checking byte " + j);
                    }
                    for (int j = offset + _buffer_size; j < buffer.Length; j++)
                    {
                        Assert.AreEqual(0, buffer[j], "Checking byte " + j);
                    }
                    offset += _buffer_size;
                    Assert.AreEqual(_workbook_size - offset,
                          stream.Available(), "offset " + offset);
                }
                Assert.AreEqual(_workbook_size % _buffer_size, stream.Available());
                Arrays.Fill(buffer, (byte)0);
                int count = stream.Read(buffer, offset,
                      _workbook_size % _buffer_size);

                Assert.AreEqual(_workbook_size % _buffer_size, count);
                for (int j = 0; j < offset; j++)
                {
                    Assert.AreEqual(0, buffer[j], "Checking byte " + j);
                }
                for (int j = offset; j < buffer.Length; j++)
                {
                    Assert.AreEqual(_workbook_data[j],
                          buffer[j], "Checking byte " + j);
                }
                Assert.AreEqual(_workbook_size, offset + count);
                for (int j = count; j < offset; j++)
                {
                    Assert.AreEqual(0, buffer[j], "byte " + j);
                }

                Assert.AreEqual(-1, stream.Read(buffer, 0, 1));
                stream.Close();
                try
                {
                    stream.Read(buffer, 0, 1);
                    Assert.Fail("Should have caught IOException");
                }
                catch (IOException)
                {
                    // as expected
                }
            }
        }

        /**
         * Test Skip
         *
         * @exception IOException
         */
        [Test]
        public void TestSkip()
        {
            DocumentInputStream[] streams = new DocumentInputStream[] {
             new DocumentInputStream(_workbook_o),
             new NDocumentInputStream(_workbook_n)
       };
            foreach (DocumentInputStream stream in streams)
            {
                Assert.AreEqual(_workbook_size, stream.Available());
                int count = stream.Available();

                while (stream.Available() >= _buffer_size)
                {
                    Assert.AreEqual(_buffer_size, stream.Skip(_buffer_size));
                    count -= _buffer_size;
                    Assert.AreEqual(count, stream.Available());
                }
                Assert.AreEqual(_workbook_size % _buffer_size,
                      stream.Skip(_buffer_size));
                Assert.AreEqual(0, stream.Available());
                stream.Reset();
                Assert.AreEqual(_workbook_size, stream.Available());
                Assert.AreEqual(_workbook_size, stream.Skip(_workbook_size * 2));
                Assert.AreEqual(0, stream.Available());
                stream.Reset();
                Assert.AreEqual(_workbook_size, stream.Available());
                Assert.AreEqual(_workbook_size,
                      stream.Skip(2 + (long)Int32.MaxValue));
                Assert.AreEqual(0, stream.Available());
            }
        }
        /**
     * Test that we can read files at multiple levels down the tree
     */
        [Test]
        public void TestReadMultipleTreeLevels()
        {
            POIDataSamples _samples = POIDataSamples.GetPublisherInstance();
            FileStream sample = _samples.GetFile("Sample.pub");

            DocumentInputStream stream;

            NPOIFSFileSystem npoifs = new NPOIFSFileSystem(sample);

            try
            {
                sample = _samples.GetFile("Sample.pub");
                OPOIFSFileSystem opoifs = new OPOIFSFileSystem(sample);

                // Ensure we have what we expect on the root
                Assert.AreEqual(npoifs, npoifs.Root.NFileSystem);
                Assert.AreEqual(npoifs, npoifs.Root.FileSystem);
                Assert.AreEqual(null, npoifs.Root.OFileSystem);
                Assert.AreEqual(null, opoifs.Root.FileSystem);
                Assert.AreEqual(opoifs, opoifs.Root.OFileSystem);
                Assert.AreEqual(null, opoifs.Root.NFileSystem);

                // Check inside
                foreach (DirectoryNode root in new DirectoryNode[] { opoifs.Root, npoifs.Root })
                {
                    // Top Level
                    Entry top = root.GetEntry("Contents");
                    Assert.AreEqual(true, top.IsDocumentEntry);
                    stream = root.CreateDocumentInputStream(top);
                    stream.Read();

                    // One Level Down
                    DirectoryNode escher = (DirectoryNode)root.GetEntry("Escher");
                    Entry one = escher.GetEntry("EscherStm");
                    Assert.AreEqual(true, one.IsDocumentEntry);
                    stream = escher.CreateDocumentInputStream(one);
                    stream.Read();

                    // Two Levels Down
                    DirectoryNode quill = (DirectoryNode)root.GetEntry("Quill");
                    DirectoryNode quillSub = (DirectoryNode)quill.GetEntry("QuillSub");
                    Entry two = quillSub.GetEntry("CONTENTS");
                    Assert.AreEqual(true, two.IsDocumentEntry);
                    stream = quillSub.CreateDocumentInputStream(two);
                    stream.Read();
                }
            }
            finally
            {
                npoifs.Close();
            }
        }
    }
}