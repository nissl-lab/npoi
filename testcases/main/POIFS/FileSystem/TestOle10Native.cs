/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
using System.IO;
using System.Collections.Generic;

using NUnit.Framework;

using TestCases;
using NPOI.POIFS.FileSystem;

namespace TestCases.POIFS.FileSystem
{
    /// <summary>
    /// Summary description for TestOle10Native
    /// </summary>
    [TestFixture]
    public class TestOle10Native
    {
        private static POIDataSamples dataSamples = POIDataSamples.GetPOIFSInstance();
        public TestOle10Native()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        [Test]
        public void TestOleNative()
        {
            POIFSFileSystem fs = new POIFSFileSystem(dataSamples.OpenResourceAsStream("oleObject1.bin"));

            Ole10Native ole = Ole10Native.CreateFromEmbeddedOleObject(fs);

            Assert.AreEqual("File1.svg", ole.GetLabel());
            Assert.AreEqual("D:\\Documents and Settings\\rsc\\My Documents\\file1.svg", ole.GetCommand());
        }

       
    }
}
