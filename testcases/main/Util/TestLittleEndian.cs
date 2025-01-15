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
using System.Diagnostics;
using NUnit.Framework;
using NPOI.Util;

namespace TestCases.Util
{
    /// <summary>
    /// Summary description for TestLittleEndian
    /// </summary>
    [TestFixture]
    public class TestLittleEndian
    {

        [Test]
        public void TestGetShort()
        {
            byte[] testdata = new byte[ LittleEndianConsts.SHORT_SIZE + 1 ];

            testdata[ 0 ] = 0x01;
            testdata[ 1 ] = unchecked(( byte ) 0xFF);
            testdata[ 2 ] = 0x02;
            short[] expected = new short[2];

            expected[ 0 ] = unchecked(( short ) 0xFF01);
            expected[ 1 ] = 0x02FF;
            Assert.AreEqual(expected[ 0 ], LittleEndian.GetShort(testdata));
            Assert.AreEqual(expected[ 1 ], LittleEndian.GetShort(testdata, 1));
        }
        [Test]
        public void TestGetUShort()
        {
            byte[] testdata = new byte[ LittleEndianConsts.SHORT_SIZE + 1 ];

            testdata[ 0 ] = 0x01;
            testdata[ 1 ] = unchecked(( byte ) 0xFF);
            testdata[ 2 ] = 0x02;

            byte[] testdata2 = new byte[ LittleEndianConsts.SHORT_SIZE + 1 ];
            
            testdata2[ 0 ] = 0x0D;
            testdata2[ 1 ] = unchecked(( byte )0x93);
            testdata2[ 2 ] = unchecked(( byte )0xFF);

            int[] expected = new int[4];

            expected[ 0 ] = 0xFF01;
            expected[ 1 ] = 0x02FF;
            expected[ 2 ] = 0x930D;
            expected[ 3 ] = 0xFF93;
            Assert.AreEqual(expected[ 0 ], LittleEndian.GetUShort(testdata));
            Assert.AreEqual(expected[ 1 ], LittleEndian.GetUShort(testdata, 1));
            Assert.AreEqual(expected[ 2 ], LittleEndian.GetUShort(testdata2));
            Assert.AreEqual(expected[ 3 ], LittleEndian.GetUShort(testdata2, 1));

            byte[] testdata3 = new byte[ LittleEndianConsts.SHORT_SIZE + 1 ];
            LittleEndian.PutShort(testdata3, 0, ( short ) expected[2] );
            LittleEndian.PutShort(testdata3, 1, ( short ) expected[3] );
            Assert.AreEqual(testdata3[ 0 ], 0x0D);
            Assert.AreEqual(testdata3[ 1 ], unchecked((byte)0x93));
            Assert.AreEqual(testdata3[ 2 ], unchecked((byte)0xFF));
            Assert.AreEqual(expected[ 2 ], LittleEndian.GetUShort(testdata3));
            Assert.AreEqual(expected[ 3 ], LittleEndian.GetUShort(testdata3, 1));
        }

        private static byte[]   _DOUBLE_array =
        {
            56, 50, unchecked((byte)-113), unchecked((byte)-4), unchecked((byte)-63), unchecked((byte)-64), unchecked((byte)-13), 63, 76, unchecked((byte)-32), unchecked((byte)-42), unchecked((byte)-35), 60, unchecked((byte)-43), 3, 64
        };
        private static byte[]   _nan_DOUBLE_array =
        {
            0x00, 0x00, 0x3C, 0x00, 0x20, 0x04, 0xFF, 0xFF
        };
        private static double[] _DOUBLEs      =
        {
            1.23456, 2.47912, Double.NaN
        };

        [Test]
        public void TestGetDouble()
        {
            Assert.AreEqual(_DOUBLEs[ 0 ], LittleEndian.GetDouble(_DOUBLE_array, 0), 0.000001 );
            Assert.AreEqual(_DOUBLEs[ 1 ], LittleEndian.GetDouble( _DOUBLE_array, LittleEndianConsts.DOUBLE_SIZE), 0.000001);
            Assert.IsTrue(Double.IsNaN(LittleEndian.GetDouble(_nan_DOUBLE_array, 0)));

            double nan = LittleEndian.GetDouble(_nan_DOUBLE_array, 0);
            byte[] data = new byte[8];
            LittleEndian.PutDouble(data, 0, nan);
            for ( int i = 0; i < data.Length; i++ )
            {
                byte b = data[i];
                Assert.AreEqual(data[i], _nan_DOUBLE_array[i]);
            }
        }

        [Test]
        public void TestGetInt()
        {
            byte[] testdata = new byte[ LittleEndianConsts.INT_SIZE + 1 ];

            testdata[ 0 ] = 0x01;
            testdata[ 1 ] = unchecked(( byte ) 0xFF);
            testdata[ 2 ] = unchecked(( byte ) 0xFF);
            testdata[ 3 ] = unchecked(( byte ) 0xFF);
            testdata[ 4 ] = 0x02;
            int[] expected = new int[2];

            expected[ 0 ] = unchecked((int)0xFFFFFF01);
            expected[ 1 ] = 0x02FFFFFF;
            Assert.AreEqual(expected[ 0 ], LittleEndian.GetInt(testdata));
            Assert.AreEqual(expected[ 1 ], LittleEndian.GetInt(testdata, 1));
        }

        [Test]
        public void TestGetLong()
        {
            byte[] testdata = new byte[ LittleEndianConsts.LONG_SIZE + 1 ];

            testdata[ 0 ] = 0x01;
            testdata[ 1 ] = unchecked(( byte ) 0xFF);
            testdata[ 2 ] = unchecked(( byte ) 0xFF);
            testdata[ 3 ] = unchecked(( byte ) 0xFF);
            testdata[ 4 ] = unchecked(( byte ) 0xFF);
            testdata[ 5 ] = unchecked(( byte ) 0xFF);
            testdata[ 6 ] = unchecked(( byte ) 0xFF);
            testdata[ 7 ] = unchecked(( byte ) 0xFF);
            testdata[ 8 ] = 0x02;
            long[] expected = new long[2];

            expected[ 0 ] = unchecked((long)0xFFFFFFFFFFFFFF01L);
            expected[ 1 ] = 0x02FFFFFFFFFFFFFFL;
            Assert.AreEqual(expected[ 0 ], LittleEndian.GetLong(testdata, 0));
            Assert.AreEqual(expected[ 1 ], LittleEndian.GetLong(testdata, 1));
        }

        [Test]
        public void TestPutShort()
        {
            byte[] expected = new byte[ LittleEndianConsts.SHORT_SIZE + 1 ];

            expected[ 0 ] = 0x01;
            expected[ 1 ] = unchecked(( byte ) 0xFF);
            expected[ 2 ] = 0x02;
            byte[] received   = new byte[ LittleEndianConsts.SHORT_SIZE + 1 ];
            short[] testdata = new short[2];

            testdata[ 0 ] = unchecked(( short ) 0xFF01);
            testdata[ 1 ] = 0x02FF;
            LittleEndian.PutShort(received, 0,  testdata[ 0 ]);
            Assert.IsTrue(ba_equivalent(received, expected, 0,
                                     LittleEndianConsts.SHORT_SIZE));
            LittleEndian.PutShort(received, 1, testdata[ 1 ]);
            Assert.IsTrue(ba_equivalent(received, expected, 1,
                                     LittleEndianConsts.SHORT_SIZE));
        }

        [Test]
        public void TestPutInt()
        {
            byte[] expected = new byte[ LittleEndianConsts.INT_SIZE + 1 ];

            expected[ 0 ] = 0x01;
            expected[ 1 ] = unchecked(( byte ) 0xFF);
            expected[ 2 ] = unchecked(( byte ) 0xFF);
            expected[ 3 ] = unchecked(( byte ) 0xFF);
            expected[ 4 ] = 0x02;
            byte[] received   = new byte[ LittleEndianConsts.INT_SIZE + 1 ];
            int[] testdata = new int[2];

            testdata[ 0 ] = unchecked((int)0xFFFFFF01);
            testdata[ 1 ] = 0x02FFFFFF;
            LittleEndian.PutInt(received, 0, testdata[ 0 ]);
            Assert.IsTrue(ba_equivalent(received, expected, 0,
                                     LittleEndianConsts.INT_SIZE));
            LittleEndian.PutInt(received, 1, testdata[ 1 ]);
            Assert.IsTrue(ba_equivalent(received, expected, 1,
                                     LittleEndianConsts.INT_SIZE));
        }

        [Test]
        public void TestPutDouble()
        {
            byte[] received = new byte[ LittleEndianConsts.DOUBLE_SIZE + 1 ];

            LittleEndian.PutDouble(received, 0,  _DOUBLEs[ 0 ]);
            Assert.IsTrue(ba_equivalent(received, _DOUBLE_array, 0,
                                     LittleEndianConsts.DOUBLE_SIZE));
            LittleEndian.PutDouble(received, 1, _DOUBLEs[ 1 ]);
            byte[] expected = new byte[ LittleEndianConsts.DOUBLE_SIZE + 1 ];

            System.Array.Copy(_DOUBLE_array, LittleEndianConsts.DOUBLE_SIZE, expected,
                             1, LittleEndianConsts.DOUBLE_SIZE);
            Assert.IsTrue(ba_equivalent(received, expected, 1,
                                     LittleEndianConsts.DOUBLE_SIZE));
        }

        [Test]
        public void TestPutLong()
        {
            byte[] expected = new byte[ LittleEndianConsts.LONG_SIZE + 1 ];

            expected[ 0 ] = 0x01;
            expected[ 1 ] = ( byte ) 0xFF;
            expected[ 2 ] = ( byte ) 0xFF;
            expected[ 3 ] = ( byte ) 0xFF;
            expected[ 4 ] = ( byte ) 0xFF;
            expected[ 5 ] = ( byte ) 0xFF;
            expected[ 6 ] = ( byte ) 0xFF;
            expected[ 7 ] = ( byte ) 0xFF;
            expected[ 8 ] = 0x02;
            byte[] received   = new byte[ LittleEndianConsts.LONG_SIZE + 1 ];
            long[] testdata = new long[2];

            testdata[ 0 ] = unchecked((long)0xFFFFFFFFFFFFFF01L);
            testdata[ 1 ] = 0x02FFFFFFFFFFFFFFL;
            LittleEndian.PutLong(received, 0, testdata[ 0 ]);
            Assert.IsTrue(ba_equivalent(received, expected, 0,
                                     LittleEndianConsts.LONG_SIZE));
            LittleEndian.PutLong(received, 1, testdata[ 1 ]);
            Assert.IsTrue(ba_equivalent(received, expected, 1,
                                     LittleEndianConsts.LONG_SIZE));
        }

        private static byte[] _good_array =
        {
            0x01, 0x02, 0x01, 0x02, 0x01, 0x02, 0x01, 0x02, 0x01, 0x02, 0x01,
            0x02, 0x01, 0x02, 0x01, 0x02
        };
        private static byte[] _bad_array  =
        {
            0x01
        };

        [Test]
        public void TestReadShort()
        {
            short       expected_value = 0x0201;
            Stream stream         = new MemoryStream(_good_array);
            int         count          = 0;

            while ((stream.Length - stream.Position) > 0)
            {
                short value = LittleEndian.ReadShort(stream);
                Assert.AreEqual(value, expected_value);
                count++;
            }
            Assert.AreEqual(count,
                         _good_array.Length / LittleEndianConsts.SHORT_SIZE);
            stream = new MemoryStream(_bad_array);
            try
            {
                LittleEndian.ReadShort(stream);
                Assert.Fail("Should have caught BufferUnderrunException");
            }
            catch (BufferUnderrunException )
            {

                // as expected
            }
        }

        [Test]
        public void TestReadInt()
        {
            int         expected_value = 0x02010201;
            Stream stream         = new MemoryStream(_good_array);
            int         count          = 0;

            while ((stream.Length - stream.Position) > 0)
            {
                int value = LittleEndian.ReadInt(stream);
                Assert.AreEqual(value, expected_value);
                count++;
            }
            Assert.AreEqual(count, _good_array.Length / LittleEndianConsts.INT_SIZE);
            stream = new MemoryStream(_bad_array);
            try
            {
                LittleEndian.ReadInt(stream);
                Assert.Fail("Should have caught BufferUnderrunException");
            }
            catch (BufferUnderrunException )
            {

                // as expected
            }
        }

        [Test]
        public void TestReadLong()
        {
            long        expected_value = 0x0201020102010201L;
            Stream stream         = new MemoryStream(_good_array);
            int         count          = 0;

            while ((stream.Length - stream.Position) > 0)
            {
                long value = LittleEndian.ReadLong(stream);
                Assert.AreEqual(value, expected_value);
                count++;
            }
            Assert.AreEqual(count,
                         _good_array.Length / LittleEndianConsts.LONG_SIZE);
            stream = new MemoryStream(_bad_array);
            try
            {
                LittleEndian.ReadLong(stream);
                Assert.Fail("Should have caught BufferUnderrunException");
            }
            catch (BufferUnderrunException )
            {

                // as expected
            }
        }

        //[Test]
        //public void TestReadFromStream()
        //{
        //    Stream stream = new MemoryStream(_good_array);
        //    byte[]      value  = LittleEndian.ReadFromStream(stream,
        //                             _good_array.Length);

        //    Assert.IsTrue(ba_equivalent(value, _good_array, 0, _good_array.Length));
        //    stream = new MemoryStream(_good_array);
        //    try
        //    {
        //        value = LittleEndian.ReadFromStream(stream,
        //                                            _good_array.Length + 1);
        //        Assert.Fail("Should have caught BufferUnderrunException");
        //    }
        //    catch (BufferUnderrunException )
        //    {

        //        // as expected
        //    }
        //}
        [Test]
        public void TestUnsignedByteToInt()
        {
            Assert.AreEqual(255, LittleEndian.UByteToInt(unchecked((byte)255)));
        }

        private bool ba_equivalent(byte [] received, byte [] expected,
                                      int offset, int size)
        {
            bool result = true;

            for (int j = offset; j < offset + size; j++)
            {
                if (received[ j ] != expected[ j ])
                {
                    Console.WriteLine("difference at index " + j);
                    result = false;
                    break;
                }
            }
            return result;
        }
        [Test]
        public void TestUnsignedShort()
        {
            Assert.AreEqual(0xffff, LittleEndian.GetUShort(new byte[] { unchecked((byte)0xff), unchecked((byte)0xff) }, 0));
        }
        
    }
}
