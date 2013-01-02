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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.Collections;
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;

    using TestCases.HSSF;
    using System.Collections.Generic;

    /**
     * 
     */
    [TestFixture]
    public class TestOLE2Embeding
    {
        [Test]
        public void TestEmbeding()
        {
            // This used to break, until bug #43116 was fixed
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("ole2-embedding.xls");

            // Check we can get at the Escher layer still
            workbook.GetAllPictures();
        }
        [Test]
        public void TestEmbeddedObjects()
        {
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("ole2-embedding.xls");

            IList<HSSFObjectData> objects = workbook.GetAllEmbeddedObjects();
            Assert.AreEqual(2, objects.Count, "Wrong number of objects");
            Assert.AreEqual("MBD06CAB431",
                ((HSSFObjectData)objects[0]).GetDirectory().Name,
                    "Wrong name for first object");
            Assert.AreEqual("MBD06CAC85A",
                    ((HSSFObjectData)
                    objects[1]).GetDirectory().Name, "Wrong name for second object");
        }
    }
}

