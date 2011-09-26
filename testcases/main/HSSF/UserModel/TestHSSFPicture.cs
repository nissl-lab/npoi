/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
namespace TestCases.HSSF.UserModel
{
    using NPOI.HSSF.UserModel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.SS.UserModel;
    using TestCases.SS.UserModel;

    /**
     * Test <c>HSSFPicture</c>.
     *
     * @author Yegor Kozlov (yegor at apache.org)
     */
    [TestClass]
    public class TestHSSFPicture:BaseTestPicture
    {
        public TestHSSFPicture()
            : base(HSSFITestDataProvider.Instance)
        {
            
        }

        [TestMethod]
        public void TestResize()
        {
            BaseTestResize(new HSSFClientAnchor(0, 0, 848, 240, (short)0, 0, (short)1, 9));
        }

        /**
         * Bug # 45829 reported ArithmeticException (/ by zero) when resizing png with zero DPI.
         */
        [TestMethod]
        public void Test45829()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sh1 = wb.CreateSheet();
            IDrawing p1 = sh1.CreateDrawingPatriarch();

            byte[] pictureData = HSSFTestDataSamples.GetTestDataFileContent("45829.png");
            int idx1 = wb.AddPicture(pictureData, PictureType.PNG);
            IPicture pic = p1.CreatePicture(new HSSFClientAnchor(), idx1);
            pic.Resize();
        }
    }
}