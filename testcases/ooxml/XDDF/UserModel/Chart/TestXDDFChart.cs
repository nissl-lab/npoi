/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.XDDF.UserModel.Chart
{
    using NPOI;
    using NPOI.OOXML;
    using NPOI.XDDF.UserModel.Chart;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    public class TestXDDFChart
    {
        [Test]
        public void TestConstruct()
        {
            // minimal test to cause ooxml-lite to include all the classes in poi-ooxml-schemas
            XDDFChart xddfChart = new XDDFChart1();

            ClassicAssert.IsNotNull(xddfChart.GetCTPlotArea());
        }
    }

    public class XDDFChart1 : XDDFChart
    {
        protected override POIXMLRelation GetChartRelation()
        {
            return null;
        }
        protected override POIXMLRelation GetChartWorkbookRelation()
        {
            return null;
        }
        protected override POIXMLFactory GetChartFactory()
        {
            return null;
        }
    };
}


