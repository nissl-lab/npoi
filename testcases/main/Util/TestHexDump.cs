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
using System.Reflection;

namespace TestCases.Util
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class TestHexDump
    {

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

            Assert.AreEqual("[]", HexDump.ToHex(new short[] { }));
            Assert.AreEqual("[000A]", HexDump.ToHex(new short[] { 0xA }));
            Assert.AreEqual(HexDump.ToHex(new short[] { 0xA, 0xB }), "[000A, 000B]");

            Assert.AreEqual("0A", HexDump.ToHex((byte)0xA));
            Assert.AreEqual("0000000A", HexDump.ToHex(0xA));

            Assert.AreEqual("[]", HexDump.ToHex(new byte[] { }));
            Assert.AreEqual("[0A]", HexDump.ToHex(new byte[] { 0xA }));
            Assert.AreEqual(HexDump.ToHex(new byte[] { 0xA, 0xB }), "[0A, 0B]");

            Assert.AreEqual(HexDump.ToHex(new byte[] { }, 10), ": 0");
            Assert.AreEqual(HexDump.ToHex(new byte[] { 0xA }, 10), "0: 0A");
            Assert.AreEqual(HexDump.ToHex(new byte[] { 0xA, 0xB }, 10), "0: 0A, 0B");
            Assert.AreEqual(HexDump.ToHex(new byte[] { 0xA, 0xB, 0xC, 0xD }, 2), "0: 0A, 0B\n2: 0C, 0D");
            Assert.AreEqual(HexDump.ToHex(new byte[] { 0xA, 0xB, 0xC, 0xD, 0xE, 0xF }, 2), "0: 0A, 0B\n2: 0C, 0D\n4: 0E, 0F");

            Assert.AreEqual("FFFF", HexDump.ToHex(unchecked((short)0xFFFF)));

            Assert.AreEqual("00000000000004D2", HexDump.ToHex(1234L));

            ConfirmStr("0xFE", HexDump.ByteToHex(-2));
            ConfirmStr("0x25", HexDump.ByteToHex(37));
            ConfirmStr("0xFFFE", HexDump.ShortToHex(-2));
            ConfirmStr("0x0005", HexDump.ShortToHex(5));
            ConfirmStr("0xFFFFFF9C", HexDump.IntToHex(-100));
            ConfirmStr("0x00001001", HexDump.IntToHex(4097));
            ConfirmStr("0xFFFFFFFFFFFF0006", HexDump.LongToHex(-65530));
            ConfirmStr("0x0000000000003FCD", HexDump.LongToHex(16333));
        }
        private static void ConfirmStr(String expected, char[] actualChars)
        {
            Assert.AreEqual(expected, new String(actualChars));
        }
        [Test]
        public void TestDumpToString()
        {
            byte[] testArray = new byte[256];

            for (int j = 0; j < 256; j++)
            {
                testArray[j] = (byte)j;
            }
            String dump = HexDump.Dump(testArray, 0, 0);
            //System.out.Println("Hex: \n" + dump);
            Assert.IsTrue(dump.Contains("0123456789:;<=>?"), "Had: \n" + dump);

            dump = HexDump.Dump(testArray, 2, 1);
            //System.out.Println("Hex: \n" + dump);
            Assert.IsTrue(dump.Contains("123456789:;<=>?@"), "Had: \n" + dump);
        }

        [Test]
        public void TestDumpToStringOutOfIndex()
        {
            byte[] testArray = new byte[0];

            try
            {
                HexDump.Dump(testArray, 0, -1);
                Assert.Fail("Should throw an exception with invalid input");
            }
            catch (IndexOutOfRangeException)
            {
                // expected
            }

            try
            {
                HexDump.Dump(testArray, 0, 1);
                Assert.Fail("Should throw an exception with invalid input");
            }
            catch (IndexOutOfRangeException)
            {
                // expected
            }
        }

        [Test]
        public void TestDumpToPrintStream()
        {
            byte[] testArray = new byte[256];

            for (int j = 0; j < 256; j++)
            {
                testArray[j] = (byte)j;
            }

            MemoryStream in1 = new MemoryStream(testArray);
            try
            {
                MemoryStream byteOut = new MemoryStream();
                BufferedStream out1 = new BufferedStream(byteOut);
                try
                {
                    HexDump.Dump(in1, out1, 0, 256);
                }
                finally
                {
                    out1.Close();
                }

                String str = Encoding.UTF8.GetString(byteOut.ToArray());
                Assert.IsTrue(str.Contains("0123456789:;<=>?"),"Had: \n" + str);
            }
            finally
            {
                in1.Close();
            }

            in1 = new MemoryStream(testArray);
            try
            {
                MemoryStream byteOut = new MemoryStream();
                BufferedStream out1 = new BufferedStream(byteOut);
                try
                {
                    // test with more than we have
                    HexDump.Dump(in1, out1, 0, 1000);
                }
                finally
                {
                    out1.Close();
                }

                String str = Encoding.UTF8.GetString(byteOut.ToArray());
                Assert.IsTrue(str.Contains("0123456789:;<=>?"), "Had: \n" + str);
            }
            finally
            {
                in1.Close();
            }

            in1 = new MemoryStream(testArray);
            try
            {
                MemoryStream byteOut = new MemoryStream();
                BufferedStream out1 = new BufferedStream(byteOut);
                try
                {
                    // test with -1
                    HexDump.Dump(in1, out1, 0, -1);
                }
                finally
                {
                    out1.Close();
                }

                String str = Encoding.UTF8.GetString(byteOut.ToArray());
                Assert.IsTrue(str.Contains("0123456789:;<=>?"), "Had: \n" + str);
            }
            finally
            {
                in1.Close();
            }

            in1 = new MemoryStream(testArray);
            try
            {
                MemoryStream byteOut = new MemoryStream();
                BufferedStream out1 = new BufferedStream(byteOut);
                try
                {
                    HexDump.Dump(in1, out1, 1, 235);
                }
                finally
                {
                    out1.Close();
                }

                String str = Encoding.UTF8.GetString(byteOut.ToArray());
                Assert.IsTrue(str.Contains("123456789:;<=>?@"), 
                    "Line contents should be Moved by one now, but Had: \n" + str);
            }
            finally
            {
                in1.Close();
            }
        }

        [Test]
        public void TestConstruct()
        {
            // to cover private constructor
            // Get the default constructor
            ConstructorInfo c = typeof(HexDump).GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null,
                new Type[] { typeof(HexDump) }, null);

            // make it callable from the outside
            //c.=(/*setter*/true);

            // call it
            //Assert.IsNotNull(c.Invoke((Object[]) null));        
        }

        [Test]
        public void TestMain()
        {
            FileInfo file = TempFile.CreateTempFile("HexDump", ".dat");
            try
            {
                FileStream out1 = new FileStream(file.FullName, FileMode.Create, FileAccess.ReadWrite);
                try
                {

                    IOUtils.Copy(new MemoryStream(Encoding.UTF8.GetBytes("teststring")), out1);
                }
                finally
                {
                    out1.Close();
                }
                Assert.IsTrue(file.Exists);
                Assert.IsTrue(file.Length > 0);

                //HexDump.Main(new String[] { file.AbsolutePath });
            }
            finally
            {
                file.Delete();
                //Assert.IsTrue(file.Exists);
            }
        }

    }
}
