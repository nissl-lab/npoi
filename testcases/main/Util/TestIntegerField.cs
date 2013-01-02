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
    /// Summary description for TestHexRead
    /// </summary>
    [TestFixture]
    public class TestIntegerField
    {
        private int[] _test_array;


        public TestIntegerField()
        {
           _test_array =new int[]
            {
                int.MinValue, -1, 0, 1, int.MaxValue
            };
        }

        [Test]
        public void TestConstructors()
        {
            try
            {
                new IntegerField(-1);
                Assert.Fail("Should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            IntegerField field = new IntegerField(2);

            Assert.AreEqual(0, field.Value);
            try
            {
                new IntegerField(-1, 1);
                Assert.Fail("Should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            field = new IntegerField(2, 0x12345678);
            Assert.AreEqual(0x12345678, field.Value);
            byte[] array = new byte[ 6 ];

            try
            {
                new IntegerField(-1, 1, array);
                Assert.Fail("Should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            field = new IntegerField(2, 0x12345678, array);
            Assert.AreEqual(0x12345678, field.Value);
            Assert.AreEqual(( byte ) 0x78, array[ 2 ]);
            Assert.AreEqual(( byte ) 0x56, array[ 3 ]);
            Assert.AreEqual(( byte ) 0x34, array[ 4 ]);
            Assert.AreEqual(( byte ) 0x12, array[ 5 ]);
            array = new byte[ 5 ];
            try
            {
                new IntegerField(2, 5, array);
                Assert.Fail("should have gotten IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            for (int j = 0; j < _test_array.Length; j++)
            {
                array = new byte[ 4 ];
                new IntegerField(0, _test_array[ j ], array);
                Assert.AreEqual(_test_array[ j ], new IntegerField(0, array).Value);
            }

            // same test as above, but using the static method
            for (int j = 0; j < _test_array.Length; j++)
            {
                array = new byte[ 4 ];
                IntegerField.Write(0, _test_array[j], array);
                Assert.AreEqual(_test_array[ j ], new IntegerField(0, array).Value);
            }
        }

        /**
         * Test set() methods
         */
        [Test]
        public void TestSet()
        {
            IntegerField field = new IntegerField(0);
            byte[]       array = new byte[ 4 ];

            for (int j = 0; j < _test_array.Length; j++)
            {
                field.Value=_test_array[ j ];
                Assert.AreEqual(_test_array[j], field.Value, "testing _1 " + j);
                field = new IntegerField(0);
                field.Set(_test_array[ j ], array);
                Assert.AreEqual(_test_array[ j ], field.Value,"testing _2 ");
                Assert.AreEqual((byte)(_test_array[j] % 256), array[0], "testing _3.0 " + _test_array[j]);
                Assert.AreEqual(( byte ) ((_test_array[ j ] >> 8) % 256),
                             array[ 1 ],"testing _3.1 " + _test_array[ j ]);
                Assert.AreEqual(( byte ) ((_test_array[ j ] >> 16) % 256),
                             array[2], "testing _3.2 " + _test_array[j]);
                Assert.AreEqual(( byte ) ((_test_array[ j ] >> 24) % 256),
                             array[3], "testing _3.3 " + _test_array[j]);
            }
        }

        /**
         * Test readFromBytes
         */
        [Test]
        public void TestReadFromBytes()
        {
            IntegerField field = new IntegerField(1);
            byte[]       array = new byte[ 4 ];

            try
            {
                field.ReadFromBytes(array);
                Assert.Fail("should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            field = new IntegerField(0);
            for (int j = 0; j < _test_array.Length; j++)
            {
                array[ 0 ] = ( byte ) (_test_array[ j ] % 256);
                array[ 1 ] = ( byte ) ((_test_array[ j ] >> 8) % 256);
                array[ 2 ] = ( byte ) ((_test_array[ j ] >> 16) % 256);
                array[ 3 ] = ( byte ) ((_test_array[ j ] >> 24) % 256);
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
            IntegerField field  = new IntegerField(0);
            byte[]       buffer = new byte[ _test_array.Length * 4 ];

            for (int j = 0; j < _test_array.Length; j++)
            {
                buffer[ (j * 4) + 0 ] = ( byte ) (_test_array[ j ] % 256);
                buffer[ (j * 4) + 1 ] = ( byte ) ((_test_array[ j ] >> 8) % 256);
                buffer[ (j * 4) + 2 ] = ( byte ) ((_test_array[ j ] >> 16) % 256);
                buffer[ (j * 4) + 3 ] = ( byte ) ((_test_array[ j ] >> 24) % 256);
            }
            MemoryStream stream = new MemoryStream(buffer);

            for (int j = 0; j < buffer.Length / 4; j++)
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
            IntegerField field = new IntegerField(0);
            byte[]       array = new byte[ 4 ];

            for (int j = 0; j < _test_array.Length; j++)
            {
                field.Value=_test_array[ j ];
                field.WriteToBytes(array);
                int val = array[ 3 ] << 24;

                val &= unchecked((int)0xFF000000);
                val += (array[ 2 ] << 16) & 0x00FF0000;
                val += (array[ 1 ] << 8) & 0x0000FF00;
                val += (array[ 0 ] & 0x000000FF);
                Assert.AreEqual(_test_array[j], val, "testing ");
            }
        }
    }
}
