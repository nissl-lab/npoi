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


namespace TestCases.Util
{

    using System;

    using NUnit.Framework;
    using System.IO;
    using NPOI.Util;

    /**
     * Tests ClassID structure.
     *
     * @author Michael Zalewski (zalewski@optonline.net)
     */
    [TestFixture]
    public class TestClassID
    {
        /**
         * Constructor
         * 
         * @param name the Test case's name
         */
        public TestClassID()
        {

        }

        /**
         * Various Tests of overridden .Equals()
         */
        [Test]
        public void TestEquals()
        {
            ClassID clsidTest1 = new ClassID(
                  new byte[] {0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
                          0x09, 0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x0f, 0x10}
                , 0
            );
            ClassID clsidTest2 = new ClassID(
                  new byte[] {0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
                          0x09, 0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x0f, 0x10}
                , 0
            );
            ClassID clsidTest3 = new ClassID(
                  new byte[] {0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
                          0x09, 0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x0f, 0x11 }
                , 0
            );
            Assert.AreEqual(clsidTest1, clsidTest1);
            Assert.AreEqual(clsidTest1, clsidTest2);
            Assert.IsFalse(clsidTest1.Equals(clsidTest3));
            Assert.IsFalse(clsidTest1.Equals(null));
        }
        /**
         * Try to Write to a buffer that is too small. This should
         *   throw an Exception
         */
        [Test]
        public void TestWriteArrayStoreException()
        {
            ClassID clsidTest = new ClassID(
                  new byte[] {0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
                          0x09, 0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x0f, 0x10}
                , 0
            );
            bool bExceptionOccurred = false;
            try
            {
                clsidTest.Write(new byte[15], 0);
            }
            catch (Exception)
            {
                bExceptionOccurred = true;
            }
            Assert.IsTrue(bExceptionOccurred);

            bExceptionOccurred = false;
            try
            {
                clsidTest.Write(new byte[16], 1);
            }
            catch (Exception)
            {
                bExceptionOccurred = true;
            }
            Assert.IsTrue(bExceptionOccurred);

            // These should work without throwing an Exception
            bExceptionOccurred = false;
            try
            {
                clsidTest.Write(new byte[16], 0);
                clsidTest.Write(new byte[17], 1);
            }
            catch (Exception)
            {
                bExceptionOccurred = true;
            }
            Assert.IsFalse(bExceptionOccurred);
        }
        /**
         * Tests the {@link PropertySet} methods. The Test file has two
         * property Set: the first one is a {@link SummaryInformation},
         * the second one is a {@link DocumentSummaryInformation}.
         */
        [Test]
        public void TestClassID1()
        {
            ClassID clsidTest = new ClassID(
                  new byte[] {0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
                          0x09, 0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x0f, 0x10}
                , 0
            );
            Assert.AreEqual(clsidTest.ToString().ToUpper(),
                                "{04030201-0605-0807-090A-0B0C0D0E0F10}"
            );
        }

    }
}