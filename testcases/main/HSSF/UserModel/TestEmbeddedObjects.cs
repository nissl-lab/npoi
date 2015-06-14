/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NUnit.Framework;

namespace TestCases.HSSF.UserModel
{
    /**
    * Tests for the embedded object fetching support in HSSF
    */
    [TestFixture]
    public class TestEmbeddedObjects
    {
        [Test]
        public void TestReadExistingObject()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            IList<HSSFObjectData> list = wb.GetAllEmbeddedObjects();
            Assert.AreEqual(list.Count, 1);
            HSSFObjectData obj = list[0];
            Assert.IsNotNull(obj.GetObjectData());
            Assert.IsNotNull(obj.GetDirectory());
            Assert.IsNotNull(obj.OLE2ClassName);
        }

        /**
         * Need to recurse into the shapes to find this one
         * See https://github.com/apache/poi/pull/2
         */
        [Test]
        public void TestReadNestedObject()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("WithCheckBoxes.xls");
            IList<HSSFObjectData> list = wb.GetAllEmbeddedObjects();
            Assert.AreEqual(list.Count, 1);
            HSSFObjectData obj = list[0];
            Assert.IsNotNull(obj.GetObjectData());
            Assert.IsNotNull(obj.OLE2ClassName);
        }

        /**
         * One with large numbers of recursivly embedded resources
         * See https://github.com/apache/poi/pull/2
         */
        [Test]
        public void TestReadManyNestedObjects()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("45538_form_Header.xls");
            IList<HSSFObjectData> list = wb.GetAllEmbeddedObjects();
            Assert.AreEqual(list.Count, 40);
        }
    }
}
