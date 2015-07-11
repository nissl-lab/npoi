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

namespace NPOI.XSSF.Model
{
    using System;

using static org.junit.Assert;



using org.apache.Commons.codec.binary;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel;
using org.junit;

    [TestFixture]
    public class TestThemesTable
{
    private String testFile = "Themes.xlsx";

    [Test]
        [Test]
    public void TestThemesTableColors(){
        XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook(testFile);
        String rgbExpected[] = {
            "ffffff", // Lt1
            "000000", // Dk1
            "eeece1", // Lt2
            "1f497d", // DK2
            "4f81bd", // Accent1
            "c0504d", // Accent2
            "9bbb59", // Accent3
            "8064a2", // Accent4
            "4bacc6", // Accent5
            "f79646", // Accent6
            "0000ff", // Hlink
            "800080"  // FolHlink
        };
        bool CreateFile = false;
        int i=0;
        foreach (Row row in workbook.GetSheetAt(0)) {
            XSSFFont font = ((XSSFRow)row).GetCell(0).CellStyle.Font;
            XSSFColor color = font.XSSFColor;
            Assert.AreEqual( rgbExpected[i], Hex.EncodeHexString(color.Rgb),"Failed color theme "+i);
            long themeIdx = font.CTFont.GetColorArray(0).Theme;
            Assert.AreEqual( i, themeIdx,"Failed color theme "+i);
            if (CreateFile) {
                XSSFCellStyle cs = (XSSFCellStyle)row.Sheet.Workbook.CreateCellStyle();
                cs.FillForegroundColor=(/*setter*/color);
                cs.FillPattern=(/*setter*/CellStyle.SOLID_FOREGROUND);
                row.CreateCell(1).CellStyle=(/*setter*/cs);
            }
            i++;
        }
        
        if (CreateFile) {
            FileOutputStream fos = new FileOutputStream("foobaa.xlsx");
            workbook.Write(fos);
            fos.Close();
        }
    }
}
