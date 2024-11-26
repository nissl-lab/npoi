/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record
{
    using System;
    using System.IO;
    using NUnit.Framework;

    using TestCases.HSSF;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    /**
     * 
     */
    [TestFixture]
    public class TestBOFRecord
    {
        [Test]
        public void TestBOFRecord1()
        {
            // This used to throw an error before - #42794
            HSSFTestDataSamples.OpenSampleFileStream("bug_42794.xls").Close();
        }
    }

}