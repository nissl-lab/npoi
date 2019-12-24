
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record
{
    using System;
    using System.Collections;
    using NUnit.Framework;
    using NPOI.HSSF.Util;
    using NPOI.HSSF.Record;
    using System.Collections.Generic;


    /**
     * Verifies that custom palette editing works correctly
     *
     * @author Brian Sanders (bsanders at risklabs dot com)
     */
    [TestFixture]
    public class TestPaletteRecord
    {
        /**
         * Tests that the default palette matches the constants of HSSFColor
         */
        [Test]
        public void TestDefaultPalette()
        {
            PaletteRecord palette = new PaletteRecord();

            //make sure all the HSSFColor constants match
            Dictionary<int, HSSFColor> colors = HSSFColor.GetIndexHash();
            foreach (KeyValuePair<int, HSSFColor> entry in colors)
            {
                int index = entry.Key;
                HSSFColor c = entry.Value;
                byte[] rgbTriplet = c.GetTriplet();
                byte[] paletteTriplet = palette.GetColor((short)index);
                String msg = "Expected HSSFColor constant to match PaletteRecord at index 0x"
                    + NPOI.Util.StringUtil.ToHexString(c.Indexed);
                Assert.AreEqual(rgbTriplet[0], paletteTriplet[0] & 0xff,msg);
                Assert.AreEqual(rgbTriplet[1], paletteTriplet[1] & 0xff,msg);
                Assert.AreEqual(rgbTriplet[2], paletteTriplet[2] & 0xff,msg);
            }
        }
    }
}