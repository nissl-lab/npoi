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

using NPOI.SS.UserModel;
using NUnit.Framework;
using NPOI.Util;
using NPOI.SS.Util;
using System.IO;
using TestCases.HSSF;
using SixLabors.ImageSharp;

namespace TestCases.SS.UserModel
{

    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class BaseTestPicture
    {
        protected ITestDataProvider _testDataProvider;
        public BaseTestPicture()
            : this(TestCases.HSSF.HSSFITestDataProvider.Instance)
        { }
        protected BaseTestPicture(ITestDataProvider testDataProvider)
        {
            _testDataProvider = testDataProvider;
        }

        public void BaseTestResize(IPicture input, IPicture Compare, double scaleX, double scaleY)
        {
            input.Resize(scaleX, scaleY);

            IClientAnchor inpCA = input.ClientAnchor;
            IClientAnchor cmpCA = Compare.ClientAnchor;

            Size inpDim = ImageUtils.GetDimensionFromAnchor(input);
            Size cmpDim = ImageUtils.GetDimensionFromAnchor(Compare);

            double emuPX = Units.EMU_PER_PIXEL;

            Assert.AreEqual(inpDim.Height, cmpDim.Height, emuPX * 6, "the image height differs");
            Assert.AreEqual(inpDim.Width, cmpDim.Width, emuPX * 6, "the image width differs");
            Assert.AreEqual(inpCA.Col1, cmpCA.Col1, "the starting column differs");
            Assert.AreEqual(inpCA.Dx1, cmpCA.Dx1, 1, "the column x-offset differs");
            Assert.AreEqual(inpCA.Dy1, cmpCA.Dy1, 1, "the column y-offset differs");
            Assert.AreEqual(inpCA.Col2, cmpCA.Col2, "the ending columns differs");
            // can't compare row heights because of variable test heights

            input.Resize();
            inpDim = ImageUtils.GetDimensionFromAnchor(input);

            Size imgDim = input.GetImageDimension();

            Assert.AreEqual(imgDim.Height, inpDim.Height / emuPX, 1, "the image height differs");
            Assert.AreEqual(imgDim.Width, inpDim.Width / emuPX, 1, "the image width differs");
        }

        [Test]
        public void TestResizeNoColumns()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();

                IRow row = sheet.CreateRow(0);

                handleResize(wb, sheet, row);
            }
            finally
            {
                //wb.Close();
            }
        }

        [Test]
        public void TestResizeWithColumns()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();

                IRow row = sheet.CreateRow(0);
                row.CreateCell(0);

                handleResize(wb, sheet, row);
            }
            finally
            {
                //wb.Close();
            }
        }


        private void handleResize(IWorkbook wb, ISheet sheet, IRow row)
        {
            IDrawing Drawing = sheet.CreateDrawingPatriarch();
            ICreationHelper CreateHelper = wb.GetCreationHelper();

            byte[] bytes = HSSFITestDataProvider.Instance.GetTestDataFileContent("logoKarmokar4.png");

            row.HeightInPoints = (/*setter*/GetImageSize(bytes).Y);

            int pictureIdx = wb.AddPicture(bytes, PictureType.PNG);

            //add a picture shape
            IClientAnchor anchor = CreateHelper.CreateClientAnchor();
            //set top-left corner of the picture,
            //subsequent call of Picture#resize() will operate relative to it
            anchor.Col1 = (/*setter*/0);
            anchor.Row1 = (/*setter*/0);

            IPicture pict = Drawing.CreatePicture(anchor, pictureIdx);

            //auto-size picture relative to its top-left corner
            pict.Resize();
        }

        private static Point GetImageSize(byte[] image)
        {
            Image img = Image.Load(new MemoryStream(image));

            Assert.IsNotNull(img);

            return new Point(img.Width, img.Height);
        }

    }
}





