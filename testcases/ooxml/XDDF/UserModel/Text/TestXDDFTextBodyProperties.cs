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

namespace TestCases.XDDF.UserModel.Text
{
    using NPOI.OpenXmlFormats.Dml;
    using NPOI.Util;
    using NPOI.XDDF.UserModel.Text;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    public class TestXDDFTextBodyProperties
    {

        [Test]
        public void TestProperties()
        {

            XDDFBodyProperties body = new XDDFTextBody(null).BodyProperties;
            CT_TextBodyProperties props = body.GetXmlObject();

            body.BottomInset = null;
            ClassicAssert.IsFalse(props.IsSetBIns());
            body.BottomInset = 3.6;
            ClassicAssert.IsTrue(props.IsSetBIns());
            ClassicAssert.AreEqual(Units.ToEMU(3.6), props.bIns);

            body.LeftInset = null;
            ClassicAssert.IsFalse(props.IsSetLIns());
            body.LeftInset = 3.6;
            ClassicAssert.IsTrue(props.IsSetLIns());
            ClassicAssert.AreEqual(Units.ToEMU(3.6), props.lIns);

            body.RightInset = null;
            ClassicAssert.IsFalse(props.IsSetRIns());
            body.RightInset = 3.6;
            ClassicAssert.IsTrue(props.IsSetRIns());
            ClassicAssert.AreEqual(Units.ToEMU(3.6), props.rIns);

            body.TopInset = null;
            ClassicAssert.IsFalse(props.IsSetTIns());
            body.TopInset = 3.6;
            ClassicAssert.IsTrue(props.IsSetTIns());
            ClassicAssert.AreEqual(Units.ToEMU(3.6), props.tIns);

            body.AutoFit = null;
            ClassicAssert.IsFalse(props.IsSetNoAutofit());
            ClassicAssert.IsFalse(props.IsSetNormAutofit());
            ClassicAssert.IsFalse(props.IsSetSpAutoFit());

            body.AutoFit = new XDDFNoAutoFit();
            ClassicAssert.IsTrue(props.IsSetNoAutofit());
            ClassicAssert.IsFalse(props.IsSetNormAutofit());
            ClassicAssert.IsFalse(props.IsSetSpAutoFit());

            body.AutoFit = new XDDFNormalAutoFit();
            ClassicAssert.IsFalse(props.IsSetNoAutofit());
            ClassicAssert.IsTrue(props.IsSetNormAutofit());
            ClassicAssert.IsFalse(props.IsSetSpAutoFit());

            body.AutoFit = new XDDFShapeAutoFit();
            ClassicAssert.IsFalse(props.IsSetNoAutofit());
            ClassicAssert.IsFalse(props.IsSetNormAutofit());
            ClassicAssert.IsTrue(props.IsSetSpAutoFit());

        }
    }
}
