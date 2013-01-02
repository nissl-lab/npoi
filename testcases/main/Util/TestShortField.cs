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
using System.IO;
using System.Collections.Generic;

using NUnit.Framework;
using NPOI.Util;

namespace TestCases.Util
{
    /// <summary>
    /// Summary description for TestShortField
    /// </summary>
    [TestFixture]
    public class TestShortField
    {
        static private short[] _test_array =
        {
            short.MinValue, ( short ) -1, ( short ) 0, ( short ) 1,
            short.MaxValue
        };

        /**
         * Test constructors.
         */
        [Test]
        public void TestConstructors()
        {
            try
            {
                new ShortField(-1);
                Assert.Fail("Should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            ShortField field = new ShortField(2);

            Assert.AreEqual(0, field.Value);
            try
            {
                new ShortField(-1, ( short ) 1);
                Assert.Fail("Should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            field = new ShortField(2, ( short ) 0x1234);
            Assert.AreEqual(0x1234, field.Value);
            byte[] array = new byte[ 4 ];

            try
            {
                new ShortField(-1, ( short ) 1, ref array);
                Assert.Fail("Should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            field = new ShortField(2, ( short ) 0x1234, ref array);
            Assert.AreEqual(( short ) 0x1234, field.Value);
            Assert.AreEqual(( byte ) 0x34, array[ 2 ]);
            Assert.AreEqual(( byte ) 0x12, array[ 3 ]);
            array = new byte[ 3 ];
            try
            {
                new ShortField(2, ( short ) 5, ref array);
                Assert.Fail("should have gotten IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            for (int j = 0; j < _test_array.Length; j++)
            {
                array = new byte[ 2 ];
                new ShortField(0, _test_array[ j ], ref array);
                Assert.AreEqual(_test_array[ j ], new ShortField(0, array).Value);
            }
        }

        /**
         * Test set() methods
         */
        [Test]
        public void TestSet()
        {
            ShortField field = new ShortField(0);
            byte[]     array = new byte[ 2 ];

            for (int j = 0; j < _test_array.Length; j++)
            {
                field.Value=_test_array[ j ];
                Assert.AreEqual(_test_array[j], field.Value, "testing _1 " + j.ToString());
                field = new ShortField(0);
                field.Set(_test_array[ j ], ref array);
                Assert.AreEqual(_test_array[ j ], field.Value,
                    "testing _2 ");
                Assert.AreEqual(( byte ) (_test_array[ j ] % 256), array[ 0 ],
                    "testing _3.0 " + _test_array[j]);
                Assert.AreEqual(( byte ) ((_test_array[ j ] >> 8) % 256),
                             array[1], "testing _3.1 " + _test_array[j]);
            }
        }

        /**
         * Test readFromBytes
         */
        [Test]
        public void TestReadFromBytes()
        {
            ShortField field = new ShortField(1);
            byte[]     array = new byte[ 2 ];

            try
            {
                field.ReadFromBytes(array);
                Assert.Fail("should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            field = new ShortField(0);
            for (int j = 0; j < _test_array.Length; j++)
            {
                array[ 0 ] = ( byte ) (_test_array[ j ] % 256);
                array[ 1 ] = ( byte ) ((_test_array[ j ] >> 8) % 256);
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
            ShortField field  = new ShortField(0);
            byte[]     buffer = new byte[ _test_array.Length * 2 ];

            for (int j = 0; j < _test_array.Length; j++)
            {
                buffer[ (j * 2) + 0 ] = ( byte ) (_test_array[ j ] % 256);
                buffer[ (j * 2) + 1 ] = ( byte ) ((_test_array[ j ] >> 8) % 256);
            }
            MemoryStream stream = new MemoryStream(buffer);

            for (int j = 0; j < buffer.Length / 2; j++)
            {
                field.ReadFromStream(stream);
                Assert.AreEqual(_test_array[ j ], field.Value,"Testing " + j);
            }
        }

        /**
         * Test writeToBytes
         */
        [Test]
        public void TestWriteToBytes()
        {
            ShortField field = new ShortField(0);
            byte[] array = new byte[2];

            for (int j = 0; j < _test_array.Length; j++)
            {
                field.Value = _test_array[j];
                field.WriteToBytes(array);
                short val = (short)(array[1] << 8);

                val &= unchecked((short)0xFF00);
                val += (short)(array[0] & 0x00FF);
                Assert.AreEqual(_test_array[j], val, "testing ");
            }
        }
    }
}
