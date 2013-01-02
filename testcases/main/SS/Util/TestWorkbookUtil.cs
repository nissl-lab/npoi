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

namespace TestCases.SS.Util
{

    using NPOI.SS.Util;
    using NUnit.Framework;
    using System;
    using System.Text;
    using NPOI.Util;

    /**
     * Tests WorkbookUtil.
     *
     * @see org.apache.poi.ss.util.WorkbookUtil
     */
    [TestFixture]
    public class TestWorkbookUtil
    {
        /**
         * borrowed test cases from 
         * {@link org.apache.poi.hssf.record.TestBoundSheetRecord#testValidNames()}
         */
        [Test]
        public void TestCreateSafeNames()
        {

            String p = "Sheet1";
            String actual = WorkbookUtil.CreateSafeSheetName(p);
            Assert.AreEqual(p, actual);

            p = "O'Brien's sales";
            actual = WorkbookUtil.CreateSafeSheetName(p);
            Assert.AreEqual(p, actual);

            p = " data # ";
            actual = WorkbookUtil.CreateSafeSheetName(p);
            Assert.AreEqual(p, actual);

            p = "data $1.00";
            actual = WorkbookUtil.CreateSafeSheetName(p);
            Assert.AreEqual(p, actual);

            // now the replaced versions ...
            actual = WorkbookUtil.CreateSafeSheetName("data?");
            Assert.AreEqual("data ", actual);

            actual = WorkbookUtil.CreateSafeSheetName("abc/def");
            Assert.AreEqual("abc def", actual);

            actual = WorkbookUtil.CreateSafeSheetName("data[0]");
            Assert.AreEqual("data 0 ", actual);

            actual = WorkbookUtil.CreateSafeSheetName("data*");
            Assert.AreEqual("data ", actual);

            actual = WorkbookUtil.CreateSafeSheetName("abc\\def");
            Assert.AreEqual("abc def", actual);

            actual = WorkbookUtil.CreateSafeSheetName("'data");
            Assert.AreEqual(" data", actual);

            actual = WorkbookUtil.CreateSafeSheetName("data'");
            Assert.AreEqual("data ", actual);

            actual = WorkbookUtil.CreateSafeSheetName("d'at'a");
            Assert.AreEqual("d'at'a", actual);

            actual = WorkbookUtil.CreateSafeSheetName(null);
            Assert.AreEqual("null", actual);

            actual = WorkbookUtil.CreateSafeSheetName("");
            Assert.AreEqual("empty", actual);

            actual = WorkbookUtil.CreateSafeSheetName("1234567890123456789012345678901TOOLONG");
            Assert.AreEqual("1234567890123456789012345678901", actual);

            actual = WorkbookUtil.CreateSafeSheetName("sheet:a4");
            Assert.AreEqual("sheet a4", actual);
        }
    }
}