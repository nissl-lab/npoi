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
using System.Collections.Generic;

using NUnit.Framework;
using System.IO;
using NPOI.Util;

namespace TestCases.Util
{
    /// <summary>
    /// Summary description for TestByteField
    /// </summary>
    [TestFixture]
    public class TestByteField
    {
        public TestByteField()
        {
            _test_array=new byte[]{
            Byte.MinValue, unchecked(( byte ) -1), ( byte ) 0, ( byte ) 1, Byte.MaxValue
            };
        }

        private byte[] _test_array;


        /**
         * Test constructors.
         */
        [Test]
        public void TestConstructors()
        {
            try
            {
                new ByteField(-1);
                Assert.Fail("Should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            ByteField field = new ByteField(2);

            Assert.AreEqual(( byte ) 0, field.Value);
            try
            {
                new ByteField(-1, ( byte ) 1);
                Assert.Fail("Should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            field = new ByteField(2, ( byte ) 3);
            Assert.AreEqual(( byte ) 3, field.Value);
            byte[] array = new byte[ 3 ];

            try
            {
                new ByteField(-1, ( byte ) 1, array);
                Assert.Fail("Should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            field = new ByteField(2, ( byte ) 4, array);
            Assert.AreEqual(( byte ) 4, field.Value);
            Assert.AreEqual(( byte ) 4, array[ 2 ]);
            array = new byte[ 2 ];
            try
            {
                new ByteField(2, ( byte ) 5, array);
                Assert.Fail("should have gotten IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            for (int j = 0; j < _test_array.Length; j++)
            {
                array = new byte[ 1 ];
                new ByteField(0, _test_array[ j ], array);
                Assert.AreEqual(_test_array[ j ], new ByteField(0, array).Value);
            }
        }

        /**
         * Test set() methods
         */
        [Test]
        public void TestSet()
        {
            ByteField field = new ByteField(0);
            byte[]    array = new byte[ 1 ];

            for (int j = 0; j < _test_array.Length; j++)
            {
                field.Value=_test_array[ j ];
                Assert.AreEqual(_test_array[j], field.Value, "testing _1 " + j);
                field = new ByteField(0);
                field.Set(_test_array[ j ], array);
                Assert.AreEqual(_test_array[j], field.Value, "testing _2 ");
                Assert.AreEqual(_test_array[j], array[0], "testing _3 ");
            }
        }

        /**
         * Test readFromBytes
         */
        [Test]
        public void TestReadFromBytes()
        {
            ByteField field = new ByteField(1);
            byte[]    array = new byte[ 1 ];

            try
            {
                field.ReadFromBytes(array);
                Assert.Fail("should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
            field = new ByteField(0);
            for (int j = 0; j < _test_array.Length; j++)
            {
                array[ 0 ] = _test_array[ j ];
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
            ByteField field  = new ByteField(0);
            byte[]    buffer = new byte[ _test_array.Length ];

            Array.Copy(_test_array, 0, buffer, 0, buffer.Length);
            MemoryStream stream = new MemoryStream(buffer);

            for (int j = 0; j < buffer.Length; j++)
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
            ByteField field = new ByteField(0);
            byte[]    array = new byte[ 1 ];

            for (int j = 0; j < _test_array.Length; j++)
            {
                field.Value=_test_array[ j ];
                field.WriteToBytes(array);
                Assert.AreEqual(_test_array[j], array[0], "testing ");
            }
        }
    }
}
