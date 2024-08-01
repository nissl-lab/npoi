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
using NPOI;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System.IO;
namespace TestCases.XSSF.Model
{

    /**
     * @author Roberto Manicardi
     */
    [TestFixture]
    public class TestMapInfo
    {

        [Test]
        public void TestMapInfoExists()
        {

            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("CustomXMLMappings.xlsx");

            MapInfo mapInfo = null;
            SingleXmlCells SingleXMLCells = null;

            foreach (POIXMLDocumentPart p in wb.GetRelations())
            {


                if (p is MapInfo)
                {
                    mapInfo = (MapInfo)p;


                    CT_MapInfo ctMapInfo = mapInfo.GetCTMapInfo();

                    Assert.IsNotNull(ctMapInfo);

                    Assert.AreEqual(1, ctMapInfo.Schema.Count);

                    foreach (XSSFMap map in mapInfo.GetAllXSSFMaps())
                    {
                        string xmlSchema = map.GetSchema();
                        Assert.IsNotNull(xmlSchema);
                    }
                }
            }

            XSSFSheet sheet1 = (XSSFSheet)wb.GetSheetAt(0);

            foreach (POIXMLDocumentPart p in sheet1.GetRelations())
            {

                if (p is SingleXmlCells)
                {
                    SingleXMLCells = (SingleXmlCells)p;
                }

            }
            Assert.IsNotNull(mapInfo);
            Assert.IsNotNull(SingleXMLCells);
        }

        [Test]
        public void TestIssue1070()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("GLIMP2.xlsx");

            ConfirmMapInfo(wb);

            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb, "GLIMP2-out");
            wb.Close();
            ConfirmMapInfo(wb2);

            wb2.Close();
        }

        private void ConfirmMapInfo(XSSFWorkbook wb)
        {
            foreach(POIXMLDocumentPart p in wb.GetRelations())
            {
                if(p is MapInfo)
                {
                    MapInfo mapInfo = (MapInfo) p;
                    CT_MapInfo ctMapInfo = mapInfo.GetCTMapInfo();

                    Assert.IsNotNull(ctMapInfo);

                    Assert.AreEqual(1, ctMapInfo.Schema.Count);
                    Assert.AreEqual("Schema1", ctMapInfo.Schema[0].ID);
                    Assert.AreEqual(1, ctMapInfo.Map.Count);
                    CT_Map map = ctMapInfo.Map[0];
                    Assert.IsNotNull(map);
                    Assert.AreEqual(1, map.ID);
                    Assert.AreEqual("DataList_Map", map.Name);
                    Assert.AreEqual("DataList", map.RootElement);
                    Assert.AreEqual("Schema1", map.SchemaID);
                    Assert.IsFalse(map.ShowImportExportValidationErrors);
                    Assert.IsTrue(map.AutoFit);
                    Assert.IsFalse(map.Append);
                    Assert.IsTrue(map.PreserveSortAFLayout);
                    Assert.IsTrue(map.PreserveFormat);

                    Assert.IsNotNull(map.DataBinding);
                    Assert.AreEqual(1, map.DataBinding.ConnectionID);
                    Assert.IsTrue(map.DataBinding.FileBinding);
                    Assert.IsNull(map.DataBinding.DataBindingName);
                    Assert.IsNull(map.DataBinding.FileBindingName);
                    Assert.AreEqual(1,map.DataBinding.DataBindingLoadMode);
                }
            }
        }
    }
}
