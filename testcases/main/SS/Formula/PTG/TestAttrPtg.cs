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

namespace TestCases.SS.Formula.PTG
{
    using System;
    using NUnit.Framework;
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;

    using TestCases.HSSF.Record;

    /**
     * Tests for {@link AttrPtg}.
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestAttrPtg : AbstractPtgTestCase
    {

        /**
         * Fix for bug visible around svn r706772.
         */
        [Test]
        public void TestReSerializeAttrChoose()
        {
            byte[] data = HexRead.ReadFromString("19, 04, 03, 00, 08, 00, 11, 00, 1A, 00, 23, 00");
            ILittleEndianInput in1 = TestcaseRecordInputStream.CreateLittleEndian(data);
            Ptg[] ptgs = Ptg.ReadTokens(data.Length, in1);
            byte[] data2 = new byte[data.Length];
            try
            {
                Ptg.SerializePtgs(ptgs, data2, 0);
            }
            catch (IndexOutOfRangeException)
            {
                throw new AssertionException("incorrect re-serialization of tAttrChoose");
            }
            Assert.IsTrue(Arrays.Equals(data, data2));
        }
    }

}