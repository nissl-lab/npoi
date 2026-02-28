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

using System;
using System.Collections.Generic;
using NPOI.Util;
namespace TestCases.SS.Util
{
    /**
  * represents one comparison Test case
  */
    public class ComparisonExample
    {
        private long _rawBitsA;
        private long _rawBitsB;
        private int _expectedResult;

        public ComparisonExample(long rawBitsA, long rawBitsB, int expectedResult)
        {
            _rawBitsA = rawBitsA;
            _rawBitsB = rawBitsB;
            _expectedResult = expectedResult;
        }

        public double GetA()
        {
            return BitConverter.Int64BitsToDouble(_rawBitsA);
        }
        public double GetB()
        {
            return BitConverter.Int64BitsToDouble(_rawBitsB);
        }
        public double GetNegA()
        {
            return -BitConverter.Int64BitsToDouble(_rawBitsA);
        }
        public double GetNegB()
        {
            return -BitConverter.Int64BitsToDouble(_rawBitsB);
        }
        public int GetExpectedResult()
        {
            return _expectedResult;
        }
    }
    /**
     * Contains specific examples of <c>double</c> value pairs and their comparison result according to Excel.
     *
     * @author Josh Micich
     */
    internal class NumberComparisonExamples
    {

        private NumberComparisonExamples()
        {
            // no instances of this class
        }



        private static ComparisonExample[] examples = InitExamples();

        private static ComparisonExample[] InitExamples()
        {

            List<ComparisonExample> temp = new List<ComparisonExample>();

            AddStepTransition(temp, 0x4010000000000005L);
            AddStepTransition(temp, 0x4010000000000010L);
            AddStepTransition(temp, 0x401000000000001CL);

            AddStepTransition(temp, 0x403CE0FFFFFFFFF1L);

            AddStepTransition(temp, 0x5010000000000006L);
            AddStepTransition(temp, 0x5010000000000010L);
            AddStepTransition(temp, 0x501000000000001AL);

            AddStepTransition(temp, 0x544CE6345CF32018L);
            AddStepTransition(temp, 0x544CE6345CF3205AL);
            AddStepTransition(temp, 0x544CE6345CF3209CL);
            AddStepTransition(temp, 0x544CE6345CF320DEL);

            AddStepTransition(temp, 0x54B250001000101DL);
            AddStepTransition(temp, 0x54B2500010001050L);
            AddStepTransition(temp, 0x54B2500010001083L);

            AddStepTransition(temp, 0x6230100010001000L);
            AddStepTransition(temp, 0x6230100010001005L);
            AddStepTransition(temp, 0x623010001000100AL);

            AddStepTransition(temp, 0x7F50300020001011L);
            AddStepTransition(temp, 0x7F5030002000102BL);
            AddStepTransition(temp, 0x7F50300020001044L);


            AddStepTransition(temp, 0x2B2BFFFF1000102AL);
            AddStepTransition(temp, 0x2B2BFFFF10001079L);
            AddStepTransition(temp, 0x2B2BFFFF100010C8L);

            AddStepTransition(temp, 0x2B2BFF001000102DL);
            AddStepTransition(temp, 0x2B2BFF0010001035L);
            AddStepTransition(temp, 0x2B2BFF001000103DL);

            AddStepTransition(temp, 0x2B61800040002024L);
            AddStepTransition(temp, 0x2B61800040002055L);
            AddStepTransition(temp, 0x2B61800040002086L);


            AddStepTransition(temp, 0x008000000000000BL);
            // just outside 'subnormal' range
            AddStepTransition(temp, 0x0010000000000007L);
            AddStepTransition(temp, 0x001000000000001BL);
            AddStepTransition(temp, 0x001000000000002FL);

            foreach (ComparisonExample ce1 in new ComparisonExample[] {
				// negative, and exponents differ by more than 1
				ce(unchecked((long)0xBF30000000000000L), unchecked((long)0xBE60000000000000L), -1),

				// negative zero *is* less than positive zero, but not easy to Get out of calculations
				ce(unchecked((long)0x0000000000000000L), unchecked((long)0x8000000000000000L), +1),
				// subnormal numbers compare without rounding for some reason
				ce(0x0000000000000000L, 0x0000000000000001L, -1),
				ce(0x0008000000000000L, 0x0008000000000001L, -1),
				ce(0x000FFFFFFFFFFFFFL, 0x000FFFFFFFFFFFFEL, +1),
				ce(0x000FFFFFFFFFFFFBL, 0x000FFFFFFFFFFFFCL, -1),
				ce(0x000FFFFFFFFFFFFBL, 0x000FFFFFFFFFFFFEL, -1),

				// across subnormal threshold (some mistakes when close)
				ce(0x000FFFFFFFFFFFFFL, 0x0010000000000000L, +1),
				ce(0x000FFFFFFFFFFFFBL, 0x0010000000000007L, +1),
				ce(0x000FFFFFFFFFFFFAL, 0x0010000000000007L, 0),

				// when a bit further apart - normal results
				ce(0x000FFFFFFFFFFFF9L, 0x0010000000000007L, -1),
				ce(0x000FFFFFFFFFFFFAL, 0x0010000000000008L, -1),
				ce(0x000FFFFFFFFFFFFBL, 0x0010000000000008L, -1),
		})
            {
                temp.Add(ce1);
            }

            ComparisonExample[] result = new ComparisonExample[temp.Count];
            temp.CopyTo(result);
            return result;
        }

        private static ComparisonExample ce(long rawBitsA, long rawBitsB, int expectedResult)
        {
            return new ComparisonExample(rawBitsA, rawBitsB, expectedResult);
        }

        private static void AddStepTransition(List<ComparisonExample> temp, long rawBits)
        {
            foreach (ComparisonExample ce1 in new ComparisonExample[] {
				ce(rawBits-1, rawBits+0, 0),
				ce(rawBits+0, rawBits+1, -1),
				ce(rawBits+1, rawBits+2, 0),
		})
            {
                temp.Add(ce1);
            }

        }

        public static ComparisonExample[] GetComparisonExamples()
        {
            ComparisonExample[] result = new ComparisonExample[examples.Length];
            Array.Copy(examples, result, examples.Length);
            return result;
        }

        public static ComparisonExample[] GetComparisonExamples2()
        {
            ComparisonExample[] result = new ComparisonExample[examples.Length];
            Array.Copy(examples, result, examples.Length);

            for (int i = 0; i < result.Length; i++)
            {
                int ha = ("a" + i).GetHashCode();
                double a = ha * Math.Pow(0.75, ha % 100);
                int hb = ("b" + i).GetHashCode();
                double b = hb * Math.Pow(0.75, hb % 100);

                result[i] = new ComparisonExample(BitConverter.DoubleToInt64Bits(a), BitConverter.DoubleToInt64Bits(b), a > b ? 1 : (a == b) ? 0 : -1);
            }

            return result;
        }
    }

}