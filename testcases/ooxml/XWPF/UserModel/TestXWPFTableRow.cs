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

namespace TestCases.XWPF.UserModel
{

    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;

    [TestFixture]
    public class TestXWPFTableRow
    {
        [Test]
        public void TestCreateRow()
        {
            CT_Row ctRow = new CT_Row();
            Assert.IsNotNull(ctRow);
        }

         [Test]
        public void TestSetGetCantSplitRow()
        {
            // create a table
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);
            // table has a single row by default; grab it
            XWPFTableRow tr = table.GetRow(0);
            Assert.IsNotNull(tr);

            tr.IsCantSplitRow = true;
            bool isCant = tr.IsCantSplitRow;

            Assert.IsTrue(isCant);
        }
         [Test]
        public void TestSetGetRepeatHeader()
        {
            // create a table
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable table = new XWPFTable(ctTable, doc);
            // table has a single row by default; grab it
            XWPFTableRow tr = table.GetRow(0);
            Assert.IsNotNull(tr);

            tr.IsRepeatHeader =true;
            bool isRpt = tr.IsRepeatHeader;
            
            Assert.IsTrue(isRpt);
        }
    }

}