﻿/* ====================================================================
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

using NPOI.XDDF.UserModel.Chart;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.XDDF.UserModel.Chart
{
    public class TestXDDFChartExtensionList
    {

        [Test]
        public void GetXmlObject()
        {
            // minimal test to include generated classes in poi-ooxml-schemas
            XDDFChartExtensionList list = new XDDFChartExtensionList();
            ClassicAssert.IsNotNull(list.GetXmlObject());

            XDDFChartExtensionList list2 = new XDDFChartExtensionList(list.GetXmlObject());
            ClassicAssert.AreEqual(list.GetXmlObject(), list2.GetXmlObject());
        }
    }
}
