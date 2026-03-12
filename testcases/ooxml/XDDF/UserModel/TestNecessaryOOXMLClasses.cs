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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.XDDF.UserModel
{


    using NPOI.OpenXmlFormats.Dml.Chart;

    using NPOI.OpenXmlFormats.Dml;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using EnumsNET;


    // aim is to Get these classes loaded and included in poi-ooxml-schemas.jar
    [TestFixture]
    public class TestNecessaryOOXMLClasses
    {

        [Test]
        public void TestProblemClasses()
        {
            CT_XYAdjustHandle ctxyAdjustHandle = new CT_XYAdjustHandle();
            ClassicAssert.IsNotNull(ctxyAdjustHandle);
            CT_PolarAdjustHandle ctPolarAdjustHandle = new CT_PolarAdjustHandle();
            ClassicAssert.IsNotNull(ctPolarAdjustHandle);
            CT_ChartLines ctChartLines = new CT_ChartLines();
            ClassicAssert.IsNotNull(ctChartLines);
            CT_DashStop ctDashStop = new CT_DashStop();
            ClassicAssert.IsNotNull(ctDashStop);
            CT_Surface ctSurface = new CT_Surface();
            ClassicAssert.IsNotNull(ctSurface);
            CT_LegendEntry ctLegendEntry = new CT_LegendEntry();
            ClassicAssert.IsNotNull(ctLegendEntry);
            CT_Shape3D ctShape3D = new CT_Shape3D();
            ClassicAssert.IsNotNull(ctShape3D);
            CT_Scene3D ctScene3D = new CT_Scene3D();
            ClassicAssert.IsNotNull(ctScene3D);
            CT_EffectContainer ctEffectContainer = new CT_EffectContainer();
            ClassicAssert.IsNotNull(ctEffectContainer);
            CT_ConnectionSite ctConnectionSite = new CT_ConnectionSite();
            ClassicAssert.IsNotNull(ctConnectionSite);
            //ST_LblAlgn stLblAlgn = ST_LblAlgn.ctr;
            //ClassicAssert.IsNotNull(stLblAlgn);
            //ST_BlackWhiteMode stBlackWhiteMode = STBlackWhiteMode.Factory.newInstance();
            //ClassicAssert.IsNotNull(stBlackWhiteMode);
            //ST_RectAlignment stRectAlignment = STRectAlignment.Factory.newInstance();
            //ClassicAssert.IsNotNull(stRectAlignment);
            //ST_TileFlipMode stTileFlipMode = STTileFlipMode.Factory.newInstance();
            //ClassicAssert.IsNotNull(stTileFlipMode);
            //ST_PresetPatternVal stPresetPatternVal = STPresetPatternVal.Factory.newInstance();
            //ClassicAssert.IsNotNull(stPresetPatternVal);
            //ST_OnOffStyleType stOnOffStyleType = STOnOffStyleType.Factory.newInstance();
            //ClassicAssert.IsNotNull(stOnOffStyleType);
            CT_LineJoinBevel ctLineJoinBevel = new CT_LineJoinBevel();
            ClassicAssert.IsNotNull(ctLineJoinBevel);
            CT_LineJoinMiterProperties ctLineJoinMiterProperties = new CT_LineJoinMiterProperties();
            ClassicAssert.IsNotNull(ctLineJoinMiterProperties);
            CT_TileInfoProperties ctTileInfoProperties = new CT_TileInfoProperties();
            ClassicAssert.IsNotNull(ctTileInfoProperties);
            CT_TableStyleTextStyle ctTableStyleTextStyle = new CT_TableStyleTextStyle();
            ClassicAssert.IsNotNull(ctTableStyleTextStyle);
            CT_HeaderFooter ctHeaderFooter = new CT_HeaderFooter();
            ClassicAssert.IsNotNull(ctHeaderFooter);
            CT_MarkerSize ctMarkerSize = new CT_MarkerSize();
            ClassicAssert.IsNotNull(ctMarkerSize);
            CT_DLbls ctdLbls = new CT_DLbls();
            ClassicAssert.IsNotNull(ctdLbls);
            CT_Marker ctMarker = new CT_Marker();
            ClassicAssert.IsNotNull(ctMarker);
            //ST_MarkerStyle stMarkerStyle = STMarkerStyle.Factory.newInstance();
            //ClassicAssert.IsNotNull(stMarkerStyle);
            CT_MarkerStyle ctMarkerStyle = new CT_MarkerStyle();
            ClassicAssert.IsNotNull(ctMarkerStyle);
            CT_ExternalData ctExternalData = new CT_ExternalData();
            ClassicAssert.IsNotNull(ctExternalData);
            CT_AxisUnit ctAxisUnit = new CT_AxisUnit();
            ClassicAssert.IsNotNull(ctAxisUnit);
            CT_LblAlgn ctLblAlgn = new CT_LblAlgn();
            ClassicAssert.IsNotNull(ctLblAlgn);
            CT_DashStopList ctDashStopList = new CT_DashStopList();
            ClassicAssert.IsNotNull(ctDashStopList);
            //ST_DispBlanksAs stDashBlanksAs = STDispBlanksAs.Factory.newInstance();
            //ClassicAssert.IsNotNull(stDashBlanksAs);
            CT_DispBlanksAs ctDashBlanksAs = new CT_DispBlanksAs();
            ClassicAssert.IsNotNull(ctDashBlanksAs);

            ST_LblAlgn e1 = Enums.Parse<ST_LblAlgn>("ctr");
            ClassicAssert.IsNotNull(e1);
            ST_BlackWhiteMode e2 = Enums.Parse<ST_BlackWhiteMode>("clr");
            ClassicAssert.IsNotNull(e2);
            ST_RectAlignment e3 = Enums.Parse<ST_RectAlignment>("ctr");
            ClassicAssert.IsNotNull(e3);
            ST_TileFlipMode e4 = Enums.Parse<ST_TileFlipMode>("xy");
            ClassicAssert.IsNotNull(e4);
            ST_PresetPatternVal e5 = Enums.Parse<ST_PresetPatternVal>("horz");
            ClassicAssert.IsNotNull(e5);
            ST_MarkerStyle e6 = Enums.Parse<ST_MarkerStyle>("circle");
            ClassicAssert.IsNotNull(e6);
            ST_DispBlanksAs e7 = Enums.Parse<ST_DispBlanksAs>("span");
            ClassicAssert.IsNotNull(e7);

            CT_TextBulletTypefaceFollowText ctTextBulletTypefaceFollowText = new CT_TextBulletTypefaceFollowText();
            ClassicAssert.IsNotNull(ctTextBulletTypefaceFollowText);
            CT_TextBulletSizeFollowText ctTextBulletSizeFollowText = new CT_TextBulletSizeFollowText();
            ClassicAssert.IsNotNull(ctTextBulletSizeFollowText);
            CT_TextBulletColorFollowText ctTextBulletColorFollowText = new CT_TextBulletColorFollowText();
            ClassicAssert.IsNotNull(ctTextBulletColorFollowText);
            CT_TextBlipBullet ctTextBlipBullet = new CT_TextBlipBullet();
            ClassicAssert.IsNotNull(ctTextBlipBullet);
        }

    }
}
