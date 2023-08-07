/*
 *  ====================================================================
 *    Licensed to the collaborators of the NPOI project under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The collaborators licenses this file to You under the Apache License, Version 2.0
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
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.SS.Formula.Functions;
using NUnit.Framework.Constraints;
using NPOI.SS.Util;
using ExtendedNumerics;
using MathNet.Numerics;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestPrecisionedDecimalHelper
    {
        public const int Precision = 15;

        private static void BigDecimalPrecisionTest(bool expected, string repre)
        {
            Assert.AreEqual(expected, DoublePrecisionHelper.IsIntegerWithDigitsDropped(
                BigDecimal.Parse(repre), Precision), repre);
        }
        private static void DoublePrecisionTest(bool expected, string repre)
        {
            Assert.AreEqual(expected, DoublePrecisionHelper.IsIntegerWithDigitsDropped(
                double.Parse(repre), Precision), repre);
        }

        [Test]
        public void IsBigDecimalAlmostInteger()
        {
            Assert.Multiple(() =>
            {
                BigDecimalPrecisionTest(true, "4.9999999999999999");
                BigDecimalPrecisionTest(true, "4.999999999999999");
                BigDecimalPrecisionTest(false, "4.99999999999999");
                BigDecimalPrecisionTest(false, "4.9999999999999");
                BigDecimalPrecisionTest(false, "12345678901234.5");
                BigDecimalPrecisionTest(false, "12345678901234.9");
                BigDecimalPrecisionTest(true, "12345678901234.99");
                BigDecimalPrecisionTest(true, "123456789012345");
                BigDecimalPrecisionTest(false, "123456789012345.25");
            });
            
        }
        [Ignore("Currently unused. Bugs to be solved.")]
        [Test]
        public void IsDoubleAlmostInteger()
        {
            Assert.Multiple(() =>
            {
                DoublePrecisionTest(true, "4.9999999999999999");
                DoublePrecisionTest(true, "4.999999999999999");
                DoublePrecisionTest(false, "4.99999999999999");
                DoublePrecisionTest(false, "4.9999999999999");
                DoublePrecisionTest(false, "12345678901234.5");
                DoublePrecisionTest(false, "12345678901234.9");
                DoublePrecisionTest(true, "12345678901234.99");
                DoublePrecisionTest(true, "123456789012345");
                DoublePrecisionTest(false, "123456789012345.25"); // TODO: bugged
            });

        }
    }
}
