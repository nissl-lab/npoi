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

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using NPOI.POIFS.Storage;
    using NPOI.POIFS.Properties;
    using TestCases.HSSF;

    /**
     * Tests for POIFSFileSystem
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestPOIFSFileSystem
    {

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
                return _is.Seek(offset,origin);
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
        [TestMethod]
        public void TestAlwaysClose()
        {

            TestIS testIS;

            // Normal case - Read until EOF and close
            testIS = new TestIS(OpenSampleStream("13224.xls"), -1);
            try
            {
                new POIFSFileSystem(testIS);
            }
            catch (IOException )
            {
                throw;
            }
            Assert.IsTrue(testIS.IsClosed(),"input stream was not closed");

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
        [TestMethod]
        public void TestShortLastBlock()
        {
            String[] files = new String[] {
			    "ShortLastBlock.qwp", "ShortLastBlock.wps"	
		    };
            POIDataSamples _samples = POIDataSamples.GetPOIFSInstance();

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

        private static Stream OpenSampleStream(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleFileStream(sampleFileName);
        }
    }
}