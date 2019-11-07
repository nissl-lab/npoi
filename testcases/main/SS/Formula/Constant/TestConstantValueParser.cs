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

namespace TestCases.SS.Formula.Constant
{
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.Util;

    using TestCases.HSSF.Record;
    using NPOI.SS.Formula.Constant;
    using NPOI.SS.UserModel;

    /**
* 
* @author Josh Micich
*/
    [TestFixture]
    public class TestConstantValueParser
    {
        private static object[] SAMPLE_VALUES = {
			true,
			null,
			1.1,
			"Sample text",
			ErrorConstant.ValueOf(FormulaError.DIV0.Code),
		};
        private static byte[] SAMPLE_ENCODING = HexRead.ReadFromString(
            "04 01 00 00 00 00 00 00 00 " +
            "00 00 00 00 00 00 00 00 00 " +
            "01 9A 99 99 99 99 99 F1 3F " +
            "02 0B 00 00 53 61 6D 70 6C 65 20 74 65 78 74 " +
            "10 07 00 00 00 00 00 00 00");
        [Test]
        public void TestGetEncodedSize()
        {
            int actual = ConstantValueParser.GetEncodedSize(SAMPLE_VALUES);
            Assert.AreEqual(51, actual);
        }
        [Test]
        public void TestEncode()
        {
            int size = ConstantValueParser.GetEncodedSize(SAMPLE_VALUES);
            byte[] data = new byte[size];

            ConstantValueParser.Encode(new LittleEndianByteArrayOutputStream(data, 0), SAMPLE_VALUES);

            if (!Arrays.Equals(data, SAMPLE_ENCODING))
            {
                Assert.Fail("Encoding differs");
            }
        }
        [Test]
        public void TestDecode()
        {
            ILittleEndianInput in1 = TestcaseRecordInputStream.CreateLittleEndian(SAMPLE_ENCODING);

            object[] values = ConstantValueParser.Parse(in1, 4);
            for (int i = 0; i < values.Length; i++)
            {
                if (!IsEqual(SAMPLE_VALUES[i], values[i]))
                {
                    Assert.Fail("Decoded result differs");
                }
            }
        }
        private static bool IsEqual(object a, object b)
        {
            if (a == null)
            {
                return b == null;
            }
            return a.Equals(b);
        }
    }

}