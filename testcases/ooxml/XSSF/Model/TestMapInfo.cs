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
using NPOI.XSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NUnit.Framework;
using System.Xml;
namespace NPOI.XSSF.Model
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
    }
}
