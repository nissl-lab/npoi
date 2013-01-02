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
    using NPOI.Util.Collections;
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;
    /**
     * Tests the implementation of the FontDetails class.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestFontDetails
    {
        private Properties properties;
        private FontDetails fontDetails;

        [SetUp]
        public void SetUp()
        {
            properties = new Properties();
            properties.Add("font.Arial.height", "13");
            properties.Add("font.Arial.characters", "a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u, v, w, x, y, z, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, ");
            properties.Add("font.Arial.widths", "6, 6, 6, 6, 6, 3, 6, 6, 3, 4, 6, 3, 9, 6, 6, 6, 6, 4, 6, 3, 6, 7, 9, 6, 5, 5, 7, 7, 7, 7, 7, 6, 8, 7, 3, 6, 7, 6, 9, 7, 8, 7, 8, 7, 7, 5, 7, 7, 9, 7, 7, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, ");
            fontDetails = FontDetails.Create("Arial", properties);

        }
        [Test]
        public void TestCreate()
        {
            Assert.AreEqual(13, fontDetails.GetHeight());
            Assert.AreEqual(6, fontDetails.GetCharWidth('a'));
            Assert.AreEqual(3, fontDetails.GetCharWidth('f'));
        }
        [Test]
        public void TestGetStringWidth()
        {
            Assert.AreEqual(9, fontDetails.GetStringWidth("af"));
        }
        [Test]
        public void TestGetCharWidth()
        {
            Assert.AreEqual(6, fontDetails.GetCharWidth('a'));
            Assert.AreEqual(9, fontDetails.GetCharWidth('='));
        }

    }
}