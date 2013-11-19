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

using NPOI.SS.Util;
using NUnit.Framework;
using System;
using System.Text;
using NPOI.Util;
namespace TestCases.SS.Util
{
    /**
     * Tests for {@link ExpandedDouble}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestExpandedDouble
    {
        private static BigInteger BIG_POW_10 = new BigInteger(1000000000);

        [Test]
        public void TestNegative()
        {
            ExpandedDouble hd = new ExpandedDouble(unchecked((long)0xC010000000000000L));

            if (hd.GetBinaryExponent() == -2046)
            {
                throw new AssertionException("identified bug - sign bit not masked out of exponent");
            }
            Assert.AreEqual(2, hd.GetBinaryExponent());
            BigInteger frac = hd.GetSignificand();
            Assert.AreEqual(64, frac.BitLength());
            Assert.AreEqual(1, frac.BitCount());
        }
        [Test]
        public void TestSubnormal()
        {
            ExpandedDouble hd = new ExpandedDouble(0x0000000000000001L);

            if (hd.GetBinaryExponent() == -1023)
            {
                throw new AssertionException("identified bug - subnormal numbers not decoded properly");
            }
            Assert.AreEqual(-1086, hd.GetBinaryExponent());
            BigInteger frac = hd.GetSignificand();
            Assert.AreEqual(64, frac.BitLength());
            Assert.AreEqual(1, frac.BitCount());
        }

        /**
         * Tests specific values for conversion from {@link ExpandedDouble} to {@link NormalisedDecimal} and back
         */
        [Test]
        public void TestRoundTripShifting()
        {
            long[] rawValues = {
				0x4010000000000004L,
				0x7010000000000004L,
				0x1010000000000004L,
				0x0010000000000001L, // near lowest normal number
				0x0010000000000000L, // lowest normal number
				0x000FFFFFFFFFFFFFL, // highest subnormal number
				0x0008000000000000L, // subnormal number

				unchecked((long)0xC010000000000004L),
				unchecked((long)0xE230100010001004L),
				0x403CE0FFFFFFFFF2L,
				0x0000000000000001L, // smallest non-zero number (subnormal)
				0x6230100010000FFEL,
				0x6230100010000FFFL,
				0x6230100010001000L,
				0x403CE0FFFFFFFFF0L, // has single digit round trip error
				0x2B2BFFFF10001079L,
		    };
            bool success = true;
            for (int i = 0; i < rawValues.Length; i++)
            {
                success &= ConfirmRoundTrip(i, rawValues[i]);
            }
            if (!success)
            {
                throw new AssertionException("One or more Test examples failed.  See stderr.");
            }
        }
        public static bool ConfirmRoundTrip(int i, long rawBitsA)
        {
            double a = BitConverter.Int64BitsToDouble(rawBitsA);
            if (a == 0.0)
            {
                // Can't represent 0.0 or -0.0 with NormalisedDecimal
                return true;
            }
            ExpandedDouble ed1;
            NormalisedDecimal nd2;
            ExpandedDouble ed3;
            try
            {
                ed1 = new ExpandedDouble(rawBitsA);
                nd2 = ed1.NormaliseBaseTen();
                CheckNormaliseBaseTenResult(ed1, nd2);

                ed3 = nd2.NormaliseBaseTwo();
            }
            catch (Exception e)
            {
                Console.WriteLine("example[" + i + "] ("
                        + FormatDoubleAsHex(a) + ") exception: " + e.Message);
                return false;
            }
            if (ed3.GetBinaryExponent() != ed1.GetBinaryExponent())
            {
                Console.WriteLine("example[" + i + "] ("
                        + FormatDoubleAsHex(a) + ") bin exp mismatch");
                return false;
            }
            BigInteger diff = ed3.GetSignificand() - (ed1.GetSignificand()).Abs();
            if (diff.Signum() == 0)
            {
                return true;
            }
            // original quantity only has 53 bits of precision
            // these quantities may have errors in the 64th bit, which hopefully don't make any difference

            if (diff.BitCount() < 2)
            {
                // errors in the 64th bit happen from time to time
                // this is well below the 53 bits of precision required
                return true;
            }

            // but bigger errors are a concern
            Console.WriteLine("example[" + i + "] ("
                    + FormatDoubleAsHex(a) + ") frac mismatch: " + diff.ToString());

            for (int j = -2; j < 3; j++)
            {
                Console.WriteLine((j < 0 ? "" : "+") + j + ": " + GetNearby(ed1, j));
            }
            for (int j = -2; j < 3; j++)
            {
                Console.WriteLine((j < 0 ? "" : "+") + j + ": " + GetNearby(nd2, j));
            }


            return false;
        }

        public static String GetBaseDecimal(ExpandedDouble hd)
        {
            /*int gg = 64 - hd.GetBinaryExponent() - 1;
            BigDecimal bd = new BigDecimal(hd.GetSignificand()).divide(new BigDecimal(BigInteger.ONE<<gg));
            int excessPrecision = bd.precision() - 23;
            if (excessPrecision > 0)
            {
                bd = bd.SetScale(bd.scale() - excessPrecision, BigDecimal.ROUND_HALF_UP);
            }
            return bd.unscaledValue().ToString();*/
            throw new NotImplementedException("This Method need BigDecimal class");
        }
        public static BigInteger GetNearby(NormalisedDecimal md, int offset)
        {
            BigInteger frac = md.ComposeFrac();
            int be = frac.BitLength() - 24 - 1;
            int sc = frac.BitLength() - 64;
            return GetNearby(frac >> (sc), be, offset);
        }

        public static BigInteger GetNearby(ExpandedDouble hd, int offset)
        {
            return GetNearby(hd.GetSignificand(), hd.GetBinaryExponent(), offset);
        }

        private static BigInteger GetNearby(BigInteger significand, int binExp, int offset)
        {
            int nExtraBits = 1;
            int nDec = (int)Math.Round(3.0 + (64 + nExtraBits) * Math.Log10(2.0));
            BigInteger newFrac = (significand << nExtraBits).Add(new BigInteger(offset));

            int gg = 64 + nExtraBits - binExp - 1;

            decimal bd = new decimal(newFrac.LongValue());
            if (gg > 0)
            {
                bd = bd/(new decimal((BigInteger.One << gg).LongValue()));
            }
            else
            {
                BigInteger frac = newFrac;
                while (frac.BitLength() + binExp < 180)
                {
                    frac = frac*(BigInteger.TEN);
                }
                int binaryExp = binExp - newFrac.BitLength() + frac.BitLength();

                bd = new decimal((frac >> (frac.BitLength() - binaryExp - 1)).LongValue());
            }
            /*int excessPrecision = bd.Precision() - nDec;
            if (excessPrecision > 0)
            {
                bd = bd.SetScale(bd.Scale() - excessPrecision, BigDecimal.ROUND_HALF_UP);
            }
            return bd.unscaledValue();*/
            throw new NotImplementedException();
        }

        private static void CheckNormaliseBaseTenResult(ExpandedDouble orig, NormalisedDecimal result)
        {
            String sigDigs = result.GetSignificantDecimalDigits();
            BigInteger frac = orig.GetSignificand();
            while (frac.BitLength() + orig.GetBinaryExponent() < 200)
            {
                frac = frac * (BIG_POW_10);
            }
            int binaryExp = orig.GetBinaryExponent() - orig.GetSignificand().BitLength();

            String origDigs = (frac << (binaryExp + 1)).ToString(10);

            if (!origDigs.StartsWith(sigDigs))
            {
                throw new AssertionException("Expected '" + origDigs + "' but got '" + sigDigs + "'.");
            }

            double dO = Double.Parse("0" + System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator + origDigs.Substring(sigDigs.Length));
            double d1 = Double.Parse(result.GetFractionalPart().ToString());
            BigInteger subDigsO = new BigInteger((int)(dO * 32768 + 0.5));
            BigInteger subDigsB = new BigInteger((int)(d1 * 32768 + 0.5));

            if (subDigsO.Equals(subDigsB))
            {
                return;
            }
            BigInteger diff = (subDigsB - subDigsO).Abs();
            if (diff.IntValue() > 100)
            {
                // 100/32768 ~= 0.003
                throw new AssertionException("minor mistake");
            }
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

