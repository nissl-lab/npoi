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
namespace TestCases.SS.Format
{
    using System;

    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NPOI.SS.Format;

    [TestFixture]
    public class TestCellFormatCondition
    {
        [Test]
        public void TestSVConditions()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            CellFormatCondition lt = CellFormatCondition.GetInstance("<", "1.5");
            ClassicAssert.IsTrue(lt.Pass(1.4));
            ClassicAssert.IsFalse(lt.Pass(1.5));
            ClassicAssert.IsFalse(lt.Pass(1.6));

            CellFormatCondition le = CellFormatCondition.GetInstance("<=", "1.5");
            ClassicAssert.IsTrue(le.Pass(1.4));
            ClassicAssert.IsTrue(le.Pass(1.5));
            ClassicAssert.IsFalse(le.Pass(1.6));

            CellFormatCondition gt = CellFormatCondition.GetInstance(">", "1.5");
            ClassicAssert.IsFalse(gt.Pass(1.4));
            ClassicAssert.IsFalse(gt.Pass(1.5));
            ClassicAssert.IsTrue(gt.Pass(1.6));

            CellFormatCondition ge = CellFormatCondition.GetInstance(">=", "1.5");
            ClassicAssert.IsFalse(ge.Pass(1.4));
            ClassicAssert.IsTrue(ge.Pass(1.5));
            ClassicAssert.IsTrue(ge.Pass(1.6));

            CellFormatCondition eqs = CellFormatCondition.GetInstance("=", "1.5");
            ClassicAssert.IsFalse(eqs.Pass(1.4));
            ClassicAssert.IsTrue(eqs.Pass(1.5));
            ClassicAssert.IsFalse(eqs.Pass(1.6));

            CellFormatCondition eql = CellFormatCondition.GetInstance("==", "1.5");
            ClassicAssert.IsFalse(eql.Pass(1.4));
            ClassicAssert.IsTrue(eql.Pass(1.5));
            ClassicAssert.IsFalse(eql.Pass(1.6));

            CellFormatCondition neo = CellFormatCondition.GetInstance("<>", "1.5");
            ClassicAssert.IsTrue(neo.Pass(1.4));
            ClassicAssert.IsFalse(neo.Pass(1.5));
            ClassicAssert.IsTrue(neo.Pass(1.6));

            CellFormatCondition nen = CellFormatCondition.GetInstance("!=", "1.5");
            ClassicAssert.IsTrue(nen.Pass(1.4));
            ClassicAssert.IsFalse(nen.Pass(1.5));
            ClassicAssert.IsTrue(nen.Pass(1.6));
        }
    }
}