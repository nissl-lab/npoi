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
    using NPOI.HSSF.UserModel;
    using TestCases.HSSF;
    using NUnit.Framework;
    using NPOI.SS.UserModel;

    /**
     * @author Josh Micich
     */
    [TestFixture]
    public class TestHSSFPatriarch
    {
        [Test]
        public void TestBasic()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet();

            IDrawing patr = sheet.CreateDrawingPatriarch();

            Assert.IsNotNull(patr);

            // assert something more interesting
        }

        // TODO - fix bug 44916 (1-May-2008)
        [Test]
        public void Test44916()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet();

            // 1. Create drawing patriarch
            IDrawing patr = sheet.CreateDrawingPatriarch();

            // 2. Try to re-get the patriarch
            IDrawing existingPatr;
            try
            {
                existingPatr = sheet.DrawingPatriarch;
            }
            catch (NullReferenceException)
            {
                throw new AssertionException("Identified bug 44916");
            }

            // 3. Use patriarch
            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 600, 245, (short)1, 1, (short)1, 2);
            anchor.AnchorType = (AnchorType)(3);
            byte[] pictureData = HSSFTestDataSamples.GetTestDataFileContent("logoKarmokar4.png");
            int idx1 = wb.AddPicture(pictureData, PictureType.PNG);
            patr.CreatePicture(anchor, idx1);

            // 4. Try to re-use patriarch later
            existingPatr = sheet.DrawingPatriarch;
            Assert.IsNotNull(existingPatr);
        }

    }
}
