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
using NUnit.Framework.Constraints;
using NPOI.SS.Util;
using ExtendedNumerics;
using MathNet.Numerics;

namespace TestCases.SS.Util
{
    [TestFixture]
    public class TestDoublePrecisionHelper
    {
        public const int Precision = 15;

        private static void BigDecimalPrecisionTest(bool expected, string repre)
        {
            Assert.AreEqual(expected, BigDecimal.Parse(repre).IsIntegerWithDigitsDropped(Precision), repre);
        }
        private static void BigDecimalPrecisionTest(bool expected, double repre)
        {
            Assert.AreEqual(expected, new BigDecimal(repre).IsIntegerWithDigitsDropped(Precision), repre.ToString());
        }
        private static void DoublePrecisionTest(bool expected, string repre)
        {
            Assert.AreEqual(expected, double.Parse(repre).IsIntegerWithDigitsDropped(Precision), repre);
        }
        private static void DoublePrecisionTest(bool expected, double repre)
        {
            Assert.AreEqual(expected, repre.IsIntegerWithDigitsDropped(Precision), repre.ToString());
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
                BigDecimalPrecisionTest(true, "123456789012345.25");

                BigDecimalPrecisionTest(true, 1 / 0.0999999999999996);
            });
        }

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
                DoublePrecisionTest(true, "123456789012345.25");

                DoublePrecisionTest(true, 1 / 0.0999999999999996);
            });

        }
    }
}
