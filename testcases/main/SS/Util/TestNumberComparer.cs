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

using TestCases.SS.Util;
using NUnit.Framework;
using System;
using NPOI.SS.Util;
using System.Text;
using NPOI.Util;
namespace TestCases.SS.Util
{
    /**
     * Tests for {@link NumberComparer}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestNumberComparer
    {

        [Test]
        public void TestAllComparisonExamples()
        {
            ComparisonExample[] examples = NumberComparisonExamples.GetComparisonExamples();
            bool success = true;

            for (int i = 0; i < examples.Length; i++)
            {
                ComparisonExample ce = examples[i];
                success &= Confirm(i, ce.GetA(), ce.GetB(), +ce.GetExpectedResult());
                success &= Confirm(i, ce.GetB(), ce.GetA(), -ce.GetExpectedResult());
                success &= Confirm(i, ce.GetNegA(), ce.GetNegB(), -ce.GetExpectedResult());
                success &= Confirm(i, ce.GetNegB(), ce.GetNegA(), +ce.GetExpectedResult());
            }
            if (!success)
            {
                throw new AssertionException("One or more cases failed.  See stderr");
            }
        }
        [Test]
        public void TestRoundTripOnComparisonExamples()
        {
            ComparisonExample[] examples = NumberComparisonExamples.GetComparisonExamples();
            bool success = true;
            for (int i = 0; i < examples.Length; i++)
            {
                ComparisonExample ce = examples[i];
                success &= ConfirmRoundTrip(i, ce.GetA());
                success &= ConfirmRoundTrip(i, ce.GetNegA());
                success &= ConfirmRoundTrip(i, ce.GetB());
                success &= ConfirmRoundTrip(i, ce.GetNegB());
            }
            if (!success)
            {
                throw new AssertionException("One or more cases failed.  See stderr");
            }

        }

        private bool ConfirmRoundTrip(int i, double a)
        {
            return TestExpandedDouble.ConfirmRoundTrip(i, BitConverter.DoubleToInt64Bits(a));
        }

        /**
         * The actual example from bug 47598
         */
        [Test]
        public void TestSpecificExampleA()
        {
            double a = 0.06 - 0.01;
            double b = 0.05;
            Assert.IsFalse(a == b);
            Assert.AreEqual(0, NumberComparer.Compare(a, b));
        }

        /**
         * The example from the nabble posting
         */
        [Test]
        public void TestSpecificExampleB()
        {
            double a = 1 + 1.0028 - 0.9973;
            double b = 1.0055;
            Assert.IsFalse(a == b);
            Assert.AreEqual(0, NumberComparer.Compare(a, b));
        }

        private static bool Confirm(int i, double a, double b, int expRes)
        {
            int actRes = NumberComparer.Compare(a, b);

            int sgnActRes = actRes < 0 ? -1 : actRes > 0 ? +1 : 0;
            if (sgnActRes != expRes)
            {
                Console.WriteLine("Mismatch example[" + i + "] ("
                        + FormatDoubleAsHex(a) + ", " + FormatDoubleAsHex(b) + ") expected "
                        + expRes + " but got " + sgnActRes);
                return false;
            }
            return true;
        }
        private static String FormatDoubleAsHex(double d)
        {
            long l = BitConverter.DoubleToInt64Bits(d);
            StringBuilder sb = new StringBuilder(20);
            sb.Append(HexDump.LongToHex(l)).Append('L');
            return sb.ToString();
        }
    }
}

