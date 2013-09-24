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
using NPOI.Util;

using TestCases.HWPF;
using NUnit.Framework;
namespace NPOI.HWPF.Model
{
    [TestFixture]
    public class TestPlexOfCps
    {
        private PlexOfCps _plexOfCps = null;
        private HWPFDocFixture _hWPFDocFixture;

        [Test]
        public void TestWriteRead()
        {
            _plexOfCps = new PlexOfCps(4);

            int last = 0;
            for (int x = 0; x < 110; x++)
            {
                byte[] intHolder = new byte[4];
                int span = (int)(110.0f * (new Random((int)DateTime.Now.Ticks).Next(0,100)/100.0));
                LittleEndian.PutInt(intHolder, span);
                _plexOfCps.AddProperty(new GenericPropertyNode(last, last + span, intHolder));
                last += span;
            }

            byte[] output = _plexOfCps.ToByteArray();
            _plexOfCps = new PlexOfCps(output, 0, output.Length, 4);
            int len = _plexOfCps.Length;
            Assert.AreEqual(len, 110);

            last = 0;
            for (int x = 0; x < len; x++)
            {
                GenericPropertyNode node = _plexOfCps.GetProperty(x);
                Assert.AreEqual(node.Start, last);
                last = node.End;
                int span = LittleEndian.GetInt(node.Bytes);
                Assert.AreEqual(node.End - node.Start, span);
            }
        }
        [SetUp]
        public void SetUp()
        {
            /**@todo verify the constructors*/
            _hWPFDocFixture = new HWPFDocFixture(this);

            _hWPFDocFixture.SetUp();
        }
        [TearDown]
        public void TearDown()
        {
            _plexOfCps = null;
            _hWPFDocFixture = null;
        }

    }
}
