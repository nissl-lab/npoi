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


using System;
using System.Collections.Generic;
using System.IO;
using NPOI.OpenXml4Net.OPC.Internal;
using NUnit.Framework;
namespace TestCases.OpenXml4Net.OPC
{

    /**
     * Test TestFileHelper class.
     *
     * @author Julien Chable
     */
    [TestFixture]
    public class TestFileHelper
    {

        /**
         * TODO - use simple JDK methods on {@link File} instead:<br/>
         * {@link File#getParentFile()} instead of {@link FileHelper#getDirectory(File)
         * {@link File#getName()} instead of {@link FileHelper#getFilename(File)
         */
        [Test]
        public void TestGetDirectory()
        {
            Dictionary<String, String> expectedValue = new Dictionary<String, String>();
            expectedValue["/dir1/test.doc"] ="/dir1";
            expectedValue["/dir1/dir2/test.doc.xml"] = "/dir1/dir2";

            foreach (String filename in expectedValue.Keys)
            {
                string f1 = expectedValue[filename];
                string f2 = FileHelper.GetDirectory(filename);

                if (false)
                {
                    // YK: The original version asserted expected values against File#getAbsolutePath():
                    Assert.IsTrue(expectedValue[filename].Equals(f2,StringComparison.InvariantCultureIgnoreCase));
                    // This comparison is platform dependent. A better approach is below
                }
                Assert.IsTrue(f1.Equals(f2));
            }
        }
    }
}