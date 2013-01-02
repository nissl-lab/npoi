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
namespace TestCases.SS.UserModel
{

    /**
     * @author Yegor Kozlov
     */
    public abstract class BaseTestPicture
    {
        protected ITestDataProvider _testDataProvider;
        protected BaseTestPicture(ITestDataProvider testDataProvider)
        {
            _testDataProvider = testDataProvider;
        }

        public void BaseTestResize(IClientAnchor referenceAnchor)
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh1 = wb.CreateSheet();
            IDrawing p1 = sh1.CreateDrawingPatriarch();
            ICreationHelper factory = wb.GetCreationHelper();

            byte[] pictureData = _testDataProvider.GetTestDataFileContent("logoKarmokar4.png");
            int idx1 = wb.AddPicture(pictureData,PictureType.PNG);
            IPicture picture = p1.CreatePicture(factory.CreateClientAnchor(), idx1);
            picture.Resize();
            IClientAnchor anchor1 = picture.GetPreferredSize();

            //assert against what would BiffViewer print if we insert the image in xls and dump the file
            Assert.AreEqual(referenceAnchor.Col1, anchor1.Col1);
            Assert.AreEqual(referenceAnchor.Row1, anchor1.Row1);
            Assert.AreEqual(referenceAnchor.Col2, anchor1.Col2);
            Assert.AreEqual(referenceAnchor.Row2, anchor1.Row2);
            Assert.AreEqual(referenceAnchor.Dx1, anchor1.Dx1);
            Assert.AreEqual(referenceAnchor.Dy1, anchor1.Dy1);
            Assert.AreEqual(referenceAnchor.Dx2, anchor1.Dx2);
            Assert.AreEqual(referenceAnchor.Dy2, anchor1.Dy2);
        }
    }
}





