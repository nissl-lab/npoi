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

namespace TestCases.HSSF.UserModel
{
    using System;
    using NPOI.DDF;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;
    /**
     * Various Tests for HSSFClientAnchor.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     * @author Yegor Kozlov (yegor at apache.org)
     */
    [TestFixture]
    public class TestHSSFClientAnchor
    {
        [Test]
        public void TestGetAnchorHeightInPoints()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)wb.CreateSheet("Test");
            HSSFClientAnchor a = new HSSFClientAnchor(0, 0, 1023, 255, (short)0, 0, (short)0, 0);
            float p = a.GetAnchorHeightInPoints(sheet);
            Assert.AreEqual(12.7, p, 0.001);

            sheet.CreateRow(0).HeightInPoints = (14);
            a = new HSSFClientAnchor(0, 0, 1023, 255, (short)0, 0, (short)0, 0);
            p = a.GetAnchorHeightInPoints(sheet);
            Assert.AreEqual(13.945, p, 0.001);

            a = new HSSFClientAnchor(0, 0, 1023, 127, (short)0, 0, (short)0, 0);
            p = a.GetAnchorHeightInPoints(sheet);
            Assert.AreEqual(6.945, p, 0.001);

            a = new HSSFClientAnchor(0, 126, 1023, 127, (short)0, 0, (short)0, 0);
            p = a.GetAnchorHeightInPoints(sheet);
            Assert.AreEqual(0.054, p, 0.001);

            a = new HSSFClientAnchor(0, 0, 1023, 0, (short)0, 0, (short)0, 1);
            p = a.GetAnchorHeightInPoints(sheet);
            Assert.AreEqual(14.0, p, 0.001);

            sheet.CreateRow(0).HeightInPoints = (12);
            a = new HSSFClientAnchor(0, 127, 1023, 127, (short)0, 0, (short)0, 1);
            p = a.GetAnchorHeightInPoints(sheet);
            Assert.AreEqual(12.372, p, 0.001);

        }

        /**
         * When HSSFClientAnchor is converted into EscherClientAnchorRecord
         * Check that dx1, dx2, dy1 and dy2 are writtem "as is".
         * (Bug 42999 reported that dx1 ans dx2 are swapped if dx1>dx2. It doesn't make sense for client anchors.)
         */
        [Test]
        public void TestConvertAnchor()
        {
            HSSFClientAnchor[] anchor = {
            new HSSFClientAnchor( 0 , 0 , 0 , 0 ,(short)0, 1,(short)1,3),
            new HSSFClientAnchor( 100 , 0 , 900 , 255 ,(short)0, 1,(short)1,3),
            new HSSFClientAnchor( 900 , 0 , 100 , 255 ,(short)0, 1,(short)1,3)
        };
            for (int i = 0; i < anchor.Length; i++)
            {
                EscherClientAnchorRecord record = (EscherClientAnchorRecord)ConvertAnchor.CreateAnchor(anchor[i]);
                Assert.AreEqual(anchor[i].Dx1, record.Dx1);
                Assert.AreEqual(anchor[i].Dx2, record.Dx2);
                Assert.AreEqual(anchor[i].Dy1, record.Dy1);
                Assert.AreEqual(anchor[i].Dy2, record.Dy2);
                Assert.AreEqual(anchor[i].Col1, record.Col1);
                Assert.AreEqual(anchor[i].Col2, record.Col2);
                Assert.AreEqual(anchor[i].Row1, record.Row1);
                Assert.AreEqual(anchor[i].Row2, record.Row2);
            }
        }

        [Test]
        public void TestAnchorHeightInPoints()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet();

            HSSFClientAnchor[] anchor = {
            new HSSFClientAnchor( 0 , 0,    0 , 0 ,(short)0, 1,(short)1, 3),
            new HSSFClientAnchor( 0 , 254 , 0 , 126 ,(short)0, 1,(short)1, 3),
            new HSSFClientAnchor( 0 , 128 , 0 , 128 ,(short)0, 1,(short)1, 3),
            new HSSFClientAnchor( 0 , 0 , 0 , 128 ,(short)0, 1,(short)1, 3),
        };
            float[] ref1 = { 25.5f, 19.125f, 25.5f, 31.875f };
            for (int i = 0; i < anchor.Length; i++)
            {
                float height = anchor[i].GetAnchorHeightInPoints(sheet);
                Assert.AreEqual(ref1[i], height, 0);
            }

        }


        /**
         * Check {@link HSSFClientAnchor} constructor does not treat 32768 as -32768.
         */
        [Test]
        public void TestCanHaveRowGreaterThan32767()
        {
            // Maximum permitted row number should be 65535.
            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short)0, 32768, (short)0, 32768);

            Assert.AreEqual(32768, anchor.Row1);
            Assert.AreEqual(32768, anchor.Row2);
        }

        /**
         * Check the maximum is not set at 255*256 instead of 256*256 - 1.
         */
        [Test]
        public void TestCanHaveRowUpTo65535()
        {
            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short)0, 65535, (short)0, 65535);

            Assert.AreEqual(65535, anchor.Row1);
            Assert.AreEqual(65535, anchor.Row2);
        }
        [Test]
        public void TestCannotHaveRowGreaterThan65535()
        {
            try
            {
                new HSSFClientAnchor(0, 0, 0, 0, (short)0, 65536, (short)0, 65536);
                Assert.Fail("Expected IllegalArgumentException to be thrown");
            }
            catch (ArgumentException)
            {
                // pass
            }
        }

        /**
         * Check the same maximum value enforced when using {@link HSSFClientAnchor#setRow1}.
         */
        [Test]
        public void TestCanSetRowUpTo65535()
        {
            HSSFClientAnchor anchor = new HSSFClientAnchor();
            anchor.Row1 = (65535);
            anchor.Row2 = (65535);

            Assert.AreEqual(65535, anchor.Row1);
            Assert.AreEqual(65535, anchor.Row2);
        }

        [Test]
        public void TestCannotSetRow1GreaterThan65535()
        {
            try
            {
                new HSSFClientAnchor().Row1 = (65536);
                Assert.Fail("Expected IllegalArgumentException to be thrown");
            }
            catch (ArgumentException)
            {
                // pass
            }
        }

        [Test]
        public void TestCannotSetRow2GreaterThan65535()
        {
            try
            {
                new HSSFClientAnchor().Row2 = (65536);
                Assert.Fail("Expected IllegalArgumentException to be thrown");
            }
            catch (ArgumentException)
            {
                // pass
            }
        }
    }
}