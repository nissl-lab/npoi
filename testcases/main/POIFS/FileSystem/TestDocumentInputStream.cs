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

using Microsoft.VisualStudio.TestTools.UnitTesting;

using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NPOI.POIFS.Storage;
using NPOI.POIFS.Properties;

namespace TestCases.POIFS.FileSystem
{
    /**
     * Class to Test POIFSDocumentReader functionality
     *
     * @author Marc Johnson
     */
    [TestClass]
    public class TestDocumentInputStream
    {

        /**
         * Constructor TestDocumentInputStream
         *
         * @param name
         *
         * @exception IOException
         */

        public TestDocumentInputStream()
        {

            int blocks = (_workbook_size + 511) / 512;

            _workbook_data = new byte[512 * blocks];

            for (int i = 0; i < _workbook_data.Length; i++)
            {
                _workbook_data[i] = unchecked((byte)-1);
            }
            for (int j = 0; j < _workbook_size; j++)
            {
                _workbook_data[j] = (byte)(j * j);
            }
            RawDataBlock[] rawBlocks = new RawDataBlock[blocks];
            MemoryStream stream =
                new MemoryStream(_workbook_data);

            for (int j = 0; j < blocks; j++)
            {
                rawBlocks[j] = new RawDataBlock(stream);
            }
            POIFSDocument document = new POIFSDocument("Workbook", rawBlocks,
                                                       _workbook_size);

            _workbook = new DocumentNode(
                document.DocumentProperty,
                new DirectoryNode(
                    new DirectoryProperty("Root Entry"), null, null));
        }

        private DocumentNode _workbook;
        private byte[] _workbook_data;
        private static int _workbook_size = 5000;

        // non-even division of _workbook_size, also non-even division of
        // any block size
        private static int _buffer_size = 6;

        /**
         * Test constructor
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestConstructor()
        {
            POIFSDocumentReader stream = new POIFSDocumentReader(_workbook);

            Assert.AreEqual(_workbook_size, stream.Available);
        }

        /**
         * Test Available behavior
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestAvailable()
        {
            POIFSDocumentReader stream = new POIFSDocumentReader(_workbook);

            Assert.AreEqual(_workbook_size, stream.Available);
            stream.Close();
            try
            {
                int i = stream.Available;
                Assert.Fail("Should have caught IOException");
            }
            catch (IOException )
            {

                // as expected
            }
        }

        /**
         * Test mark/reset/markSupported.
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestMarkFunctions()
        {
            POIFSDocumentReader stream = new POIFSDocumentReader(_workbook);
            byte[] buffer = new byte[_workbook_size / 5];

            stream.Read(buffer);
            for (int j = 0; j < buffer.Length; j++)
            {
                Assert.AreEqual(_workbook_data[j],
                             buffer[j], "checking byte " + j);
            }
            Assert.AreEqual(_workbook_size - buffer.Length, stream.Available);
            stream.Seek(0, SeekOrigin.Begin);
            Assert.AreEqual(_workbook_size, stream.Available);
            stream.Read(buffer);
            //stream.Seek(12, SeekOrigin.Begin);
            //mark(12)
            stream.Read(buffer);
            Assert.AreEqual(_workbook_size - (2 * buffer.Length),
                         stream.Available);
            for (int j = buffer.Length; j < (2 * buffer.Length); j++)
            {
                Assert.AreEqual(_workbook_data[j],
                             buffer[j - buffer.Length], "checking byte " + j);
            }
            //reset()
            stream.Seek(-buffer.Length, SeekOrigin.Current);
            Assert.AreEqual(_workbook_size - buffer.Length, stream.Available);
            stream.Read(buffer);
            Assert.AreEqual(_workbook_size - (2 * buffer.Length),
                         stream.Available);
            for (int j = buffer.Length; j < (2 * buffer.Length); j++)
            {
                Assert.AreEqual(_workbook_data[j],
                             buffer[j - buffer.Length], "checking byte " + j);
            }
            //Assert.IsTrue(stream.markSupported());
        }

        /**
         * Test simple Read method
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestReadSingleByte()
        {
            POIFSDocumentReader stream = new POIFSDocumentReader(_workbook);
            int remaining = _workbook_size;

            for (int j = 0; j < _workbook_size; j++)
            {
                int b = stream.ReadByte();
                Assert.IsTrue(b >= 0, "checking sign of " + j);
                Assert.AreEqual(_workbook_data[j],
                             (byte)b, "validating byte " + j);
                remaining--;
                Assert.AreEqual(
                             remaining, stream.Available, "checking remaining after Reading byte " + j);
            }
            Assert.AreEqual(-1, stream.ReadByte());
            stream.Close();
            try
            {
                stream.ReadByte();
                Assert.Fail("Should have caught IOException");
            }
            catch (IOException )
            {

                // as expected
            }
        }

        /**
         * Test buffered Read
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestBufferRead()
        {
            POIFSDocumentReader stream = new POIFSDocumentReader(_workbook);

            try
            {
                stream.Read(null);
                Assert.Fail("Should have caught NullPointerException");
            }
            catch (NullReferenceException )
            {

                // as expected
            }

            // Test Reading zero Length buffer
            Assert.AreEqual(0, stream.Read(new byte[0]));
            Assert.AreEqual(_workbook_size, stream.Available);
            byte[] buffer = new byte[_buffer_size];
            int offset = 0;

            while (stream.Available >= buffer.Length)
            {
                Assert.AreEqual(_buffer_size, stream.Read(buffer));
                for (int j = 0; j < buffer.Length; j++)
                {
                    Assert.AreEqual(
                                 _workbook_data[offset], buffer[j], "in main loop, byte " + offset);
                    offset++;
                }
                Assert.AreEqual(_workbook_size - offset,
                             stream.Available, "offset " + offset);
            }
            Assert.AreEqual(_workbook_size % _buffer_size, stream.Available);
            for (int i = 0; i < buffer.Length; i++)
            {
                buffer[i] = (byte)0;
            }
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
                Assert.AreEqual(0, buffer[j], "checking remainder, byte " + j);
            }
            Assert.AreEqual(-1, stream.Read(buffer));
            stream.Close();
            try
            {
                stream.Read(buffer);
                Assert.Fail("Should have caught IOException");
            }
            catch (IOException )
            {

                // as expected
            }
        }

        /**
         * Test complex buffered Read
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestComplexBufferRead()
        {
            POIFSDocumentReader stream = new POIFSDocumentReader(_workbook);

            try
            {
                stream.Read(null, 0, 1);
                Assert.Fail("Should have caught NullPointerException");
            }
            catch (NullReferenceException )
            {

                // as expected
            }

            // Test illegal offsets and Lengths
            try
            {
                stream.Read(new byte[5], -4, 0);
                Assert.Fail("Should have caught IndexOutOfBoundsException");
            }
            catch (IndexOutOfRangeException )
            {

                // as expected
            }
            try
            {
                stream.Read(new byte[5], 0, -4);
                Assert.Fail("Should have caught IndexOutOfBoundsException");
            }
            catch (IndexOutOfRangeException )
            {

                // as expected
            }
            try
            {
                stream.Read(new byte[5], 0, 6);
                Assert.Fail("Should have caught IndexOutOfBoundsException");
            }
            catch (IndexOutOfRangeException )
            {

                // as expected
            }

            // Test Reading zero
            Assert.AreEqual(0, stream.Read(new byte[5], 0, 0));
            Assert.AreEqual(_workbook_size, stream.Available);
            byte[] buffer = new byte[_workbook_size];
            int offset = 0;

            while (stream.Available >= _buffer_size)
            {
                for (int i = 0; i < buffer.Length; i++)
                {
                    buffer[i] = (byte)0;
                }
                Assert.AreEqual(_buffer_size,
                             stream.Read(buffer, offset, _buffer_size));
                for (int j = 0; j < offset; j++)
                {
                    Assert.AreEqual(0, buffer[j], "checking byte " + j);
                }
                for (int j = offset; j < (offset + _buffer_size); j++)
                {
                    Assert.AreEqual(_workbook_data[j],
                                 buffer[j], "checking byte " + j);
                }
                for (int j = offset + _buffer_size; j < buffer.Length; j++)
                {
                    Assert.AreEqual(0, buffer[j], "checking byte " + j);
                }
                offset += _buffer_size;
                Assert.AreEqual(_workbook_size - offset,
                             stream.Available, "offset " + offset);
            }
            Assert.AreEqual(_workbook_size % _buffer_size, stream.Available);

            for (int i = 0; i < buffer.Length; i++)
            {
                buffer[i] = (byte)0;
            }
            int count = stream.Read(buffer, offset,
                                    _workbook_size % _buffer_size);

            Assert.AreEqual(_workbook_size % _buffer_size, count);
            for (int j = 0; j < offset; j++)
            {
                Assert.AreEqual(0, buffer[j], "checking byte " + j);
            }
            for (int j = offset; j < buffer.Length; j++)
            {
                Assert.AreEqual(_workbook_data[j],
                             buffer[j], "checking byte " + j);
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
            catch (IOException )
            {

                // as expected
            }
        }

        /**
         * Test Skip
         *
         * @exception IOException
         */
        [TestMethod]
        public void TestSkip()
        {
            POIFSDocumentReader stream = new POIFSDocumentReader(_workbook);

            Assert.AreEqual(_workbook_size, stream.Available);
            int count = stream.Available;

            while (stream.Available >= _buffer_size)
            {
                Assert.AreEqual(_buffer_size, stream.Skip(_buffer_size));
                count -= _buffer_size;
                Assert.AreEqual(count, stream.Available);
            }
            Assert.AreEqual(_workbook_size % _buffer_size,
                         stream.Skip(_buffer_size));
            Assert.AreEqual(0, stream.Available);
            stream.Seek(0, SeekOrigin.Begin);
            Assert.AreEqual(_workbook_size, stream.Available);
            Assert.AreEqual(_workbook_size, stream.Skip(_workbook_size * 2));
            Assert.AreEqual(0, stream.Available);
            stream.Seek(0, SeekOrigin.Begin);
            Assert.AreEqual(_workbook_size, stream.Available);
            Assert.AreEqual(_workbook_size,
                         stream.Skip(2 + (long)int.MaxValue));
            Assert.AreEqual(0, stream.Available);
        }
    }
}