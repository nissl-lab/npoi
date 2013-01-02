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
    /// Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class TestHexDump
    {
        public TestHexDump()
        {
            //
            // TODO: Add constructor logic here
            //
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

        private char ToHex(int n)
        {
            char[] hexChars =
                {
                    '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C',
                    'D', 'E', 'F'
                };

            return hexChars[n % 16];
        }
        private char toAscii(int c)
        {
            char rval = '.';

            if ((c >= 32) && (c <= 126))
            {
                rval = (char)c;
            }
            return rval;
        }

        [Test]
        public void TestDump()
        {
            byte[] testArray = new byte[256];

            for (int j = 0; j < 256; j++)
            {
                testArray[j] = (byte)j;
            }
            MemoryStream stream = new MemoryStream();

            HexDump.Dump(testArray, 0, stream, 0);
            byte[] outputArray = new byte[16 * (73 + HexDump.EOL.Length)];

            for (int j = 0; j < 16; j++)
            {
                int offset = (73 + HexDump.EOL.Length) * j;

                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)ToHex(j);
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)' ';
                for (int k = 0; k < 16; k++)
                {
                    outputArray[offset++] = (byte)ToHex(j);
                    outputArray[offset++] = (byte)ToHex(k);
                    outputArray[offset++] = (byte)' ';
                }
                for (int k = 0; k < 16; k++)
                {
                    outputArray[offset++] = (byte)toAscii((j * 16) + k);
                }
                System.Array.Copy(Encoding.UTF8.GetBytes(HexDump.EOL), 0, outputArray, offset,
                                 Encoding.UTF8.GetBytes(HexDump.EOL).Length);
            }
            byte[] actualOutput = stream.ToArray();

            Assert.AreEqual(outputArray.Length,
                         actualOutput.Length, "array size mismatch");

            for (int j = 0; j < outputArray.Length; j++)
            {
                Assert.AreEqual(outputArray[j],
                             actualOutput[j], "array[ " + j + "] mismatch");
            }

            // verify proper behavior with non-zero offset
            stream = new MemoryStream();
            HexDump.Dump(testArray, 0x10000000, stream, 0);
            outputArray = new byte[16 * (73 + HexDump.EOL.Length)];
            for (int j = 0; j < 16; j++)
            {
                int offset = (73 + HexDump.EOL.Length) * j;

                outputArray[offset++] = (byte)'1';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)ToHex(j);
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)' ';
                for (int k = 0; k < 16; k++)
                {
                    outputArray[offset++] = (byte)ToHex(j);
                    outputArray[offset++] = (byte)ToHex(k);
                    outputArray[offset++] = (byte)' ';
                }
                for (int k = 0; k < 16; k++)
                {
                    outputArray[offset++] = (byte)toAscii((j * 16) + k);
                }
                System.Array.Copy(Encoding.UTF8.GetBytes(HexDump.EOL), 0, outputArray, offset,
                                 Encoding.UTF8.GetBytes(HexDump.EOL).Length);
            }
            actualOutput = stream.ToArray();
            Assert.AreEqual(outputArray.Length,
                         actualOutput.Length, "array size mismatch");
            for (int j = 0; j < outputArray.Length; j++)
            {
                Assert.AreEqual(outputArray[j], actualOutput[j], "array[ " + j + "] mismatch");
            }

            // verify proper behavior with negative offset
            stream = new MemoryStream();
            HexDump.Dump(testArray, 0xFF000000, stream, 0);
            outputArray = new byte[16 * (73 + HexDump.EOL.Length)];
            for (int j = 0; j < 16; j++)
            {
                int offset = (73 + HexDump.EOL.Length) * j;

                outputArray[offset++] = (byte)'F';
                outputArray[offset++] = (byte)'F';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)ToHex(j);
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)' ';
                for (int k = 0; k < 16; k++)
                {
                    outputArray[offset++] = (byte)ToHex(j);
                    outputArray[offset++] = (byte)ToHex(k);
                    outputArray[offset++] = (byte)' ';
                }
                for (int k = 0; k < 16; k++)
                {
                    outputArray[offset++] = (byte)toAscii((j * 16) + k);
                }
                System.Array.Copy(Encoding.UTF8.GetBytes(HexDump.EOL), 0, outputArray, offset,
                                 Encoding.UTF8.GetBytes(HexDump.EOL).Length);
            }
            actualOutput = stream.ToArray();
            Assert.AreEqual(outputArray.Length, actualOutput.Length, "array size mismatch");
            for (int j = 0; j < outputArray.Length; j++)
            {
                Assert.AreEqual(outputArray[j],
                             actualOutput[j], "array[ " + j + "] mismatch");
            }

            // verify proper behavior with non-zero index
            stream = new MemoryStream();
            HexDump.Dump(testArray, 0x10000000, stream, 0x81);
            outputArray = new byte[(8 * (73 + HexDump.EOL.Length)) - 1];
            for (int j = 0; j < 8; j++)
            {
                int offset = (73 + HexDump.EOL.Length) * j;

                outputArray[offset++] = (byte)'1';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)'0';
                outputArray[offset++] = (byte)ToHex(j + 8);
                outputArray[offset++] = (byte)'1';
                outputArray[offset++] = (byte)' ';
                for (int k = 0; k < 16; k++)
                {
                    int index = 0x81 + (j * 16) + k;

                    if (index < 0x100)
                    {
                        outputArray[offset++] = (byte)ToHex(index / 16);
                        outputArray[offset++] = (byte)ToHex(index);
                    }
                    else
                    {
                        outputArray[offset++] = (byte)' ';
                        outputArray[offset++] = (byte)' ';
                    }
                    outputArray[offset++] = (byte)' ';
                }
                for (int k = 0; k < 16; k++)
                {
                    int index = 0x81 + (j * 16) + k;

                    if (index < 0x100)
                    {
                        outputArray[offset++] = (byte)toAscii(index);
                    }
                }
                System.Array.Copy(Encoding.UTF8.GetBytes(HexDump.EOL), 0, outputArray, offset,
                                 Encoding.UTF8.GetBytes(HexDump.EOL).Length);
            }
            actualOutput = stream.ToArray();
            Assert.AreEqual(outputArray.Length,
                         actualOutput.Length, "array size mismatch");
            for (int j = 0; j < outputArray.Length; j++)
            {
                Assert.AreEqual(outputArray[j],
                             actualOutput[j], "array[ " + j + "] mismatch");
            }

            // verify proper behavior with negative index
            try
            {
                HexDump.Dump(testArray, 0x10000000, new MemoryStream(),
                             -1);
                Assert.Fail("should have caught ArrayIndexOutOfBoundsException on negative index");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }

            // verify proper behavior with index that is too large
            try
            {
                HexDump.Dump(testArray, 0x10000000, new MemoryStream(),
                             testArray.Length);
                Assert.Fail("should have caught ArrayIndexOutOfBoundsException on large index");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }

            // verify proper behaviour with empty byte array
            MemoryStream os = new MemoryStream();
            HexDump.Dump(new byte[0], 0, os, 0);
            Assert.AreEqual("No Data" + Environment.NewLine, Encoding.UTF8.GetString(os.ToArray()));

        }

        [Test]
        public void TestToHex()
        {
            Assert.AreEqual("000A", HexDump.ToHex((short)0xA));
            Assert.AreEqual("0A", HexDump.ToHex((byte)0xA));
            Assert.AreEqual("0000000A", HexDump.ToHex(0xA));
            //FFFF is out of short(Java) range
            Assert.AreEqual("7FFF", HexDump.ToHex((short)0x7FFF));
        }
    }
}
