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

using NUnit.Framework;

using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NPOI.POIFS.Storage;
using NPOI.POIFS.Properties;


namespace TestCases.POIFS.FileSystem
{
    /**
     * Class to Test POIFSDocumentWriter functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestDocumentOutputStream
    {

        /**
         * Constructor TestDocumentOutputStream
         *
         * @param name
         *
         * @exception IOException
         */

        public TestDocumentOutputStream()
        {

        }

        /**
         * Test Write(int) behavior
         *
         * @exception IOException
         */
        [Test]
        public void TestWrite1()
        {
            MemoryStream stream = new MemoryStream();
            DocumentOutputStream dstream = new DocumentOutputStream(stream, 25);

            for (int j = 0; j < 25; j++)
            {
                dstream.Write(j);
            }
            try
            {
                dstream.Write(0);
                Assert.Fail("Should have caught IOException");
            }
            catch (IOException)
            {
            }
            byte[] output = stream.ToArray();

            Assert.AreEqual(25, output.Length);
            for (int j = 0; j < 25; j++)
            {
                Assert.AreEqual((byte)j, output[j]);
            }
            stream.Close();
        }

        /**
         * Test Write(byte[]) behavior
         *
         * @exception IOException
         */
        [Test]
        public void TestWrite2()
        {
            MemoryStream stream = new MemoryStream();
            DocumentOutputStream dstream = new DocumentOutputStream(stream, 25);

            for (int j = 0; j < 6; j++)
            {
                byte[] array = new byte[4];
                for (int i = 0; i < array.Length; i++)
                {
                    array[i] = (byte)j;
                }
                dstream.Write(array);
            }
            try
            {
                byte[] array = new byte[4];
                for (int i = 0; i < array.Length; i++)
                {
                    array[i] = (byte)7;
                }

                dstream.Write(array);
                Assert.Fail("Should have caught IOException");
            }
            catch (IOException )
            {
            }
            byte[] output = stream.ToArray();

            Assert.AreEqual(24, output.Length);
            for (int j = 0; j < 6; j++)
            {
                for (int k = 0; k < 4; k++)
                {
                    Assert.AreEqual((byte)j,
                                 output[(j * 4) + k], ((j * 4) + k).ToString());
                }
            }
            stream.Close();
        }

        /**
         * Test Write(byte[], int, int) behavior
         *
         * @exception IOException
         */
        [Test]
        public void TestWrite3()
        {
            MemoryStream stream = new MemoryStream();
            DocumentOutputStream dstream = new DocumentOutputStream(stream, 25);
            byte[] array = new byte[50];

            for (int j = 0; j < 50; j++)
            {
                array[j] = (byte)j;
            }
            dstream.Write(array, 1, 25);
            try
            {
                dstream.Write(array, 0, 1);
                Assert.Fail("Should have caught IOException");
            }
            catch (IOException )
            {
            }
            byte[] output = stream.ToArray();

            Assert.AreEqual(25, output.Length);
            for (int j = 0; j < 25; j++)
            {
                Assert.AreEqual((byte)(j + 1), output[j]);
            }
            stream.Close();
        }

        /**
         * Test WriteFiller()
         *
         * @exception IOException
         */
        [Test]
        public void TestWriteFiller()
        {
            MemoryStream stream = new MemoryStream();
            DocumentOutputStream dstream = new DocumentOutputStream(stream, 25);

            for (int j = 0; j < 25; j++)
            {
                dstream.Write(j);
            }
            try
            {
                dstream.Write(0);
                Assert.Fail("Should have caught IOException");
            }
            catch (IOException )
            {
            }
            dstream.WriteFiller(100, (byte)0xff);
            byte[] output = stream.ToArray();

            Assert.AreEqual(100, output.Length);
            for (int j = 0; j < 25; j++)
            {
                Assert.AreEqual((byte)j, output[j]);
            }
            for (int j = 25; j < 100; j++)
            {
                Assert.AreEqual((byte)0xff, output[j], j.ToString());
            }
            stream.Close();
        }

    }
}