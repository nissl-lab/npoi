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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.Util;

namespace TestCases.Util
{
    /// <summary>
    /// Summary description for TestLittleEndian
    /// </summary>
    [TestClass]
    public class TestLittleEndian
    {
        public TestLittleEndian()
        {

        }

        private TestContext TestContextInstance;

        /// <summary>
        ///Gets or sets the Test context which provides
        ///information about and functionality for the current Test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return TestContextInstance;
            }
            set
            {
                TestContextInstance = value;
            }
        }
        [TestMethod]
        public void TestGetShort()
        {
            byte[] testdata = new byte[ LittleEndianConstants.SHORT_SIZE + 1 ];

            testdata[ 0 ] = 0x01;
            testdata[ 1 ] = unchecked(( byte ) 0xFF);
            testdata[ 2 ] = 0x02;
            short[] expected = new short[2];

            expected[ 0 ] = unchecked(( short ) 0xFF01);
            expected[ 1 ] = 0x02FF;
            Assert.AreEqual(expected[ 0 ], LittleEndian.GetShort(testdata));
            Assert.AreEqual(expected[ 1 ], LittleEndian.GetShort(testdata, 1));
        }
        [TestMethod]
        public void TestGetUShort()
        {
            byte[] testdata = new byte[ LittleEndianConstants.SHORT_SIZE + 1 ];

            testdata[ 0 ] = 0x01;
            testdata[ 1 ] = unchecked(( byte ) 0xFF);
            testdata[ 2 ] = 0x02;

            byte[] testdata2 = new byte[ LittleEndianConstants.SHORT_SIZE + 1 ];
            
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

            byte[] testdata3 = new byte[ LittleEndianConstants.SHORT_SIZE + 1 ];
            LittleEndian.PutShort(testdata3, 0, ( short ) expected[2] );
            LittleEndian.PutShort(testdata3, 1, ( short ) expected[3] );
            Assert.AreEqual(testdata3[ 0 ], 0x0D);
            Assert.AreEqual(testdata3[ 1 ], unchecked((byte)0x93));
            Assert.AreEqual(testdata3[ 2 ], unchecked((byte)0xFF));
            Assert.AreEqual(expected[ 2 ], LittleEndian.GetUShort(testdata3));
            Assert.AreEqual(expected[ 3 ], LittleEndian.GetUShort(testdata3, 1));
            //System.out.println("TD[1][0]: "+LittleEndian.GetUShort(testdata)+" expecting 65281");
            //System.out.println("TD[1][1]: "+LittleEndian.GetUShort(testdata, 1)+" expecting 767");
            //System.out.println("TD[2][0]: "+LittleEndian.GetUShort(testdata2)+" expecting 37645");
            //System.out.println("TD[2][1]: "+LittleEndian.GetUShort(testdata2, 1)+" expecting 65427");
            //System.out.println("TD[3][0]: "+LittleEndian.GetUShort(testdata3)+" expecting 37645");
            //System.out.println("TD[3][1]: "+LittleEndian.GetUShort(testdata3, 1)+" expecting 65427");
            
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

        [TestMethod]
        public void TestGetDouble()
        {
            Assert.AreEqual(_DOUBLEs[ 0 ], LittleEndian.GetDouble(_DOUBLE_array), 0.000001 );
            Assert.AreEqual(_DOUBLEs[ 1 ], LittleEndian.GetDouble( _DOUBLE_array, LittleEndianConstants.DOUBLE_SIZE), 0.000001);
            Assert.IsTrue(Double.IsNaN(LittleEndian.GetDouble(_nan_DOUBLE_array)));

            double nan = LittleEndian.GetDouble(_nan_DOUBLE_array);
            byte[] data = new byte[8];
            LittleEndian.PutDouble(data, nan);
            for ( int i = 0; i < data.Length; i++ )
            {
                byte b = data[i];
                Assert.AreEqual(data[i], _nan_DOUBLE_array[i]);
            }
        }

        [TestMethod]
        public void TestGetInt()
        {
            byte[] testdata = new byte[ LittleEndianConstants.INT_SIZE + 1 ];

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

        [TestMethod]
        public void TestGetLong()
        {
            byte[] testdata = new byte[ LittleEndianConstants.LONG_SIZE + 1 ];

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
            Assert.AreEqual(expected[ 0 ], LittleEndian.GetLong(testdata));
            Assert.AreEqual(expected[ 1 ], LittleEndian.GetLong(testdata, 1));
        }

        [TestMethod]
        public void TestPutShort()
        {
            byte[] expected = new byte[ LittleEndianConstants.SHORT_SIZE + 1 ];

            expected[ 0 ] = 0x01;
            expected[ 1 ] = unchecked(( byte ) 0xFF);
            expected[ 2 ] = 0x02;
            byte[] received   = new byte[ LittleEndianConstants.SHORT_SIZE + 1 ];
            short[] testdata = new short[2];

            testdata[ 0 ] = unchecked(( short ) 0xFF01);
            testdata[ 1 ] = 0x02FF;
            LittleEndian.PutShort(received, testdata[ 0 ]);
            Assert.IsTrue(ba_equivalent(received, expected, 0,
                                     LittleEndianConstants.SHORT_SIZE));
            LittleEndian.PutShort(received, 1, testdata[ 1 ]);
            Assert.IsTrue(ba_equivalent(received, expected, 1,
                                     LittleEndianConstants.SHORT_SIZE));
        }

        [TestMethod]
        public void TestPutInt()
        {
            byte[] expected = new byte[ LittleEndianConstants.INT_SIZE + 1 ];

            expected[ 0 ] = 0x01;
            expected[ 1 ] = unchecked(( byte ) 0xFF);
            expected[ 2 ] = unchecked(( byte ) 0xFF);
            expected[ 3 ] = unchecked(( byte ) 0xFF);
            expected[ 4 ] = 0x02;
            byte[] received   = new byte[ LittleEndianConstants.INT_SIZE + 1 ];
            int[] testdata = new int[2];

            testdata[ 0 ] = unchecked((int)0xFFFFFF01);
            testdata[ 1 ] = 0x02FFFFFF;
            LittleEndian.PutInt(received, testdata[ 0 ]);
            Assert.IsTrue(ba_equivalent(received, expected, 0,
                                     LittleEndianConstants.INT_SIZE));
            LittleEndian.PutInt(received, 1, testdata[ 1 ]);
            Assert.IsTrue(ba_equivalent(received, expected, 1,
                                     LittleEndianConstants.INT_SIZE));
        }

        [TestMethod]
        public void TestPutDouble()
        {
            byte[] received = new byte[ LittleEndianConstants.DOUBLE_SIZE + 1 ];

            LittleEndian.PutDouble(received, _DOUBLEs[ 0 ]);
            Assert.IsTrue(ba_equivalent(received, _DOUBLE_array, 0,
                                     LittleEndianConstants.DOUBLE_SIZE));
            LittleEndian.PutDouble(received, 1, _DOUBLEs[ 1 ]);
            byte[] expected = new byte[ LittleEndianConstants.DOUBLE_SIZE + 1 ];

            System.Array.Copy(_DOUBLE_array, LittleEndianConstants.DOUBLE_SIZE, expected,
                             1, LittleEndianConstants.DOUBLE_SIZE);
            Assert.IsTrue(ba_equivalent(received, expected, 1,
                                     LittleEndianConstants.DOUBLE_SIZE));
        }

        [TestMethod]
        public void TestPutLong()
        {
            byte[] expected = new byte[ LittleEndianConstants.LONG_SIZE + 1 ];

            expected[ 0 ] = 0x01;
            expected[ 1 ] = ( byte ) 0xFF;
            expected[ 2 ] = ( byte ) 0xFF;
            expected[ 3 ] = ( byte ) 0xFF;
            expected[ 4 ] = ( byte ) 0xFF;
            expected[ 5 ] = ( byte ) 0xFF;
            expected[ 6 ] = ( byte ) 0xFF;
            expected[ 7 ] = ( byte ) 0xFF;
            expected[ 8 ] = 0x02;
            byte[] received   = new byte[ LittleEndianConstants.LONG_SIZE + 1 ];
            long[] testdata = new long[2];

            testdata[ 0 ] = unchecked((long)0xFFFFFFFFFFFFFF01L);
            testdata[ 1 ] = 0x02FFFFFFFFFFFFFFL;
            LittleEndian.PutLong(received, testdata[ 0 ]);
            Assert.IsTrue(ba_equivalent(received, expected, 0,
                                     LittleEndianConstants.LONG_SIZE));
            LittleEndian.PutLong(received, 1, testdata[ 1 ]);
            Assert.IsTrue(ba_equivalent(received, expected, 1,
                                     LittleEndianConstants.LONG_SIZE));
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

        [TestMethod]
        public void TestReadShort()
        {
            short       expected_value = 0x0201;
            Stream stream         = new MemoryStream(_good_array);
            int         count          = 0;

            while (true)
            {
                short value = LittleEndian.ReadShort(stream);

                if (value == 0)
                {
                    break;
                }
                Assert.AreEqual(value, expected_value);
                count++;
            }
            Assert.AreEqual(count,
                         _good_array.Length / LittleEndianConstants.SHORT_SIZE);
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

        [TestMethod]
        public void TestReadInt()
        {
            int         expected_value = 0x02010201;
            Stream stream         = new MemoryStream(_good_array);
            int         count          = 0;

            while (true)
            {
                int value = LittleEndian.ReadInt(stream);

                if (value == 0)
                {
                    break;
                }
                Assert.AreEqual(value, expected_value);
                count++;
            }
            Assert.AreEqual(count, _good_array.Length / LittleEndianConstants.INT_SIZE);
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

        [TestMethod]
        public void TestReadLong()
        {
            long        expected_value = 0x0201020102010201L;
            Stream stream         = new MemoryStream(_good_array);
            int         count          = 0;

            while (true)
            {
                long value = LittleEndian.ReadLong(stream);

                if (value == 0)
                {
                    break;
                }
                Assert.AreEqual(value, expected_value);
                count++;
            }
            Assert.AreEqual(count,
                         _good_array.Length / LittleEndianConstants.LONG_SIZE);
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

        [TestMethod]
        public void TestReadFromStream()
        {
            Stream stream = new MemoryStream(_good_array);
            byte[]      value  = LittleEndian.ReadFromStream(stream,
                                     _good_array.Length);

            Assert.IsTrue(ba_equivalent(value, _good_array, 0, _good_array.Length));
            stream = new MemoryStream(_good_array);
            try
            {
                value = LittleEndian.ReadFromStream(stream,
                                                    _good_array.Length + 1);
                Assert.Fail("Should have caught BufferUnderrunException");
            }
            catch (BufferUnderrunException )
            {

                // as expected
            }
        }
        [TestMethod]
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
        [TestMethod]
        public void TestUnsignedShort()
        {
            Assert.AreEqual(0xffff, LittleEndian.GetUShort(new byte[] { unchecked((byte)0xff), unchecked((byte)0xff) }, 0));
        }
        [TestMethod]
        public void TestEffective()
        {
            byte[] data = new byte[2048];
            //Stopwatch w1 = new Stopwatch();
            //w1.Start();
            //for (int t = 0; t < 10000;t++ )
                for (int i = 0; i > -1024; i--)
                    LittleEndian.PutShort(data, Math.Abs(i) * 2, (short)(i - 1));
                for (int j = 0; j > -1024; j--)
                    Assert.AreEqual((short)(j - 1), LittleEndian.GetShort(data, Math.Abs(j) * 2));
            //w1.Stop();
            //System.Console.WriteLine(w1.ElapsedMilliseconds);


            Stopwatch w2 = new Stopwatch();
            w2.Start();
            for (int t = 0; t < 10000; t++)
                for (int i = 0; i < 1024; i++)
                    LittleEndian.PutUShort2(data, i * 2, i + 1);
            w2.Stop();
            System.Console.WriteLine(w2.ElapsedMilliseconds);
        }
    }
}
