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
namespace TestCases.HWPF.Converter
{
    using NPOI.HWPF;
    using NPOI.HWPF.UserModel;
    using System;
    
    using NPOI.HWPF.Converter;
    using NUnit.Framework;
    /**
     * Test cases for {@link AbstractWordUtils}
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */
    [TestFixture]
    public class AbstractWordUtilsTest
    {
        /**
         * Test case for {@link AbstractWordUtils#buildTableCellEdgesArray(Table)}
         */
        [Test]
        public void TestBuildTableCellEdgesArray()
        {
            HWPFDocument document = HWPFTestDataSamples
                    .OpenSampleFile("table-merges.doc");
            Range range = document.GetRange();
            Table table = range.GetTable(range.GetParagraph(0));

            int[] result = AbstractWordUtils.BuildTableCellEdgesArray(table);
            Assert.AreEqual(6, result.Length);

            Assert.AreEqual(0000, result[0]);
            Assert.AreEqual(1062, result[1]);
            Assert.AreEqual(5738, result[2]);
            Assert.AreEqual(6872, result[3]);
            Assert.AreEqual(8148, result[4]);
            Assert.AreEqual(9302, result[5]);
        }
    }
}
