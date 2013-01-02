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
using System.Text;
using System.Collections.Generic;
using System.IO;


using NUnit.Framework;
using NPOI.Util;

namespace TestCases.Util
{
    /// <summary>
    /// Summary description for TestLongField
    /// </summary>
    [TestFixture]
    public class TestLongField
    {
        public TestLongField()
        {
            _test_array =new long[]
            {
                long.MinValue, -1L, 0L, 1L, long.MaxValue
            };
        }

        private long[] _test_array;

        /**
         * Test constructors.
         */
        [Test]
        public void TestConstructors()
        {
            try
            {
                new LongField(-1);
                Assert.Fail("Should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            LongField field = new LongField(2);

            Assert.AreEqual(0L, field.Value);
            try
            {
                new LongField(-1, 1L);
                Assert.Fail("Should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            field = new LongField(2, 0x123456789ABCDEF0L);
            Assert.AreEqual(0x123456789ABCDEF0L, field.Value);
            byte[] array = new byte[ 10 ];

            try
            {
                new LongField(-1, 1L, array);
                Assert.Fail("Should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            field = new LongField(2, 0x123456789ABCDEF0L, array);
            Assert.AreEqual(0x123456789ABCDEF0L, field.Value);
            Assert.AreEqual(( byte ) 0xF0, array[ 2 ]);
            Assert.AreEqual(( byte ) 0xDE, array[ 3 ]);
            Assert.AreEqual(( byte ) 0xBC, array[ 4 ]);
            Assert.AreEqual(( byte ) 0x9A, array[ 5 ]);
            Assert.AreEqual(( byte ) 0x78, array[ 6 ]);
            Assert.AreEqual(( byte ) 0x56, array[ 7 ]);
            Assert.AreEqual(( byte ) 0x34, array[ 8 ]);
            Assert.AreEqual(( byte ) 0x12, array[ 9 ]);
            array = new byte[ 9 ];
            try
            {
                new LongField(2, 5L, array);
                Assert.Fail("should have gotten IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            for (int j = 0; j < _test_array.Length; j++)
            {
                array = new byte[ 8 ];
                new LongField(0, _test_array[ j ], array);
                Assert.AreEqual(_test_array[ j ], new LongField(0, array).Value);
            }
        }

        /**
         * Test set() methods
         */
        [Test]
        public void TestSet()
        {
            LongField field = new LongField(0);
            byte[]    array = new byte[ 8 ];

            for (int j = 0; j < _test_array.Length; j++)
            {
                field.Value=_test_array[ j ];
                Assert.AreEqual(_test_array[j], field.Value, "testing _1 " + j);
                field = new LongField(0);
                field.Set(_test_array[ j ],array);
                Assert.AreEqual(_test_array[ j ], field.Value,"testing _2 ");
                Assert.AreEqual(( byte ) (_test_array[ j ] % 256), array[ 0 ],"testing _3.0 " + _test_array[ j ]);
                Assert.AreEqual(( byte ) ((_test_array[ j ] >> 8) % 256),
                             array[ 1 ],"testing _3.1 " + _test_array[ j ]);
                Assert.AreEqual(( byte ) ((_test_array[ j ] >> 16) % 256),
                             array[ 2 ],"testing _3.2 " + _test_array[ j ]);
                Assert.AreEqual(( byte ) ((_test_array[ j ] >> 24) % 256),
                             array[ 3 ],"testing _3.3 " + _test_array[ j ]);
                Assert.AreEqual(( byte ) ((_test_array[ j ] >> 32) % 256),array[ 4 ],"testing _3.4 " + _test_array[ j ]);
                Assert.AreEqual(( byte ) ((_test_array[ j ] >> 40) % 256),
                             array[ 5 ],"testing _3.5 " + _test_array[ j ]);
                Assert.AreEqual(( byte ) ((_test_array[ j ] >> 48) % 256),
                             array[6], "testing _3.6 " + _test_array[j]);
                Assert.AreEqual(( byte ) ((_test_array[ j ] >> 56) % 256),
                             array[7], "testing _3.7 " + _test_array[j]);
            }
        }

        /**
         * Test readFromBytes
         */
        [Test]
        public void TestReadFromBytes()
        {
            LongField field = new LongField(1);
            byte[]    array = new byte[ 8 ];

            try
            {
                field.ReadFromBytes(array);
                Assert.Fail("should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            field = new LongField(0);
            for (int j = 0; j < _test_array.Length; j++)
            {
                array[ 0 ] = ( byte ) (_test_array[ j ] % 256);
                array[ 1 ] = ( byte ) ((_test_array[ j ] >> 8) % 256);
                array[ 2 ] = ( byte ) ((_test_array[ j ] >> 16) % 256);
                array[ 3 ] = ( byte ) ((_test_array[ j ] >> 24) % 256);
                array[ 4 ] = ( byte ) ((_test_array[ j ] >> 32) % 256);
                array[ 5 ] = ( byte ) ((_test_array[ j ] >> 40) % 256);
                array[ 6 ] = ( byte ) ((_test_array[ j ] >> 48) % 256);
                array[ 7 ] = ( byte ) ((_test_array[ j ] >> 56) % 256);
                field.ReadFromBytes(array);
                Assert.AreEqual(_test_array[j], field.Value, "testing " + j);
            }
        }

        /**
         * Test readFromStream
         *
         * @exception IOException
         */
        [Test]
        public void TestReadFromStream()
        {
            LongField field  = new LongField(0);
            byte[]    buffer = new byte[ _test_array.Length * 8 ];

            for (int j = 0; j < _test_array.Length; j++)
            {
                buffer[ (j * 8) + 0 ] = ( byte ) (_test_array[ j ] % 256);
                buffer[ (j * 8) + 1 ] = ( byte ) ((_test_array[ j ] >> 8) % 256);
                buffer[ (j * 8) + 2 ] = ( byte ) ((_test_array[ j ] >> 16) % 256);
                buffer[ (j * 8) + 3 ] = ( byte ) ((_test_array[ j ] >> 24) % 256);
                buffer[ (j * 8) + 4 ] = ( byte ) ((_test_array[ j ] >> 32) % 256);
                buffer[ (j * 8) + 5 ] = ( byte ) ((_test_array[ j ] >> 40) % 256);
                buffer[ (j * 8) + 6 ] = ( byte ) ((_test_array[ j ] >> 48) % 256);
                buffer[ (j * 8) + 7 ] = ( byte ) ((_test_array[ j ] >> 56) % 256);
            }
            MemoryStream stream = new MemoryStream(buffer);

            for (int j = 0; j < buffer.Length / 8; j++)
            {
                field.ReadFromStream(stream);
                Assert.AreEqual(_test_array[j], field.Value, "Testing " + j);
            }
        }

        /**
         * Test writeToBytes
         */
        [Test]
        public void TestWriteToBytes()
        {
            LongField field = new LongField(0);
            byte[]    array = new byte[ 8 ];

            for (int j = 0; j < _test_array.Length; j++)
            {
                field.Value=_test_array[ j ];
                field.WriteToBytes(array);
                long val = (( long ) array[ 7 ]) << 56;

                val &= unchecked((long)0xFF00000000000000L);
                val += ((( long ) array[ 6 ]) << 48) & 0x00FF000000000000L;
                val += ((( long ) array[ 5 ]) << 40) & 0x0000FF0000000000L;
                val += ((( long ) array[ 4 ]) << 32) & 0x000000FF00000000L;
                val += ((( long ) array[ 3 ]) << 24) & 0x00000000FF000000L;
                val += ((( long ) array[ 2 ]) << 16) & 0x0000000000FF0000L;
                val += ((( long ) array[ 1 ]) << 8) & 0x000000000000FF00L;
                val += (array[ 0 ] & 0x00000000000000FFL);
                Assert.AreEqual(_test_array[ j ], val,"testing ");
            }
        }
    }
}
