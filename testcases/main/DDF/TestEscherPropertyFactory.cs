
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

namespace TestCases.DDF
{

    using System;
    using System.Text;
    using System.Collections;
    using System.IO;

    using NUnit.Framework;
    using NPOI.DDF;
    using NPOI.Util;

    /**
     * @author Glen Stampoultzis  (glens @ superlinksoftware.com)
     */
    [TestFixture]
    public class TestEscherPropertyFactory
    {
        [Test]
        public void TestCreateProperties()
        {
            String dataStr = "41 C1 " +     // propid, complex ind
                    "03 00 00 00 " +         // size of complex property
                    "01 00 " +              // propid, complex ind
                    "00 00 00 00 " +         // value
                    "41 C1 " +              // propid, complex ind
                    "03 00 00 00 " +         // size of complex property
                    "01 02 03 " +
                    "01 02 03 "
                    ;
            byte[] data = HexRead.ReadFromString(dataStr);
            EscherPropertyFactory f = new EscherPropertyFactory();
            IList props = f.CreateProperties(data, 0, (short)3);
            EscherComplexProperty p1 = (EscherComplexProperty)props[0];
            Assert.AreEqual(unchecked((short)0xC141), p1.Id);
            Assert.AreEqual("[01, 02, 03]", HexDump.ToHex(p1.ComplexData));

            EscherComplexProperty p3 = (EscherComplexProperty)props[2];
            Assert.AreEqual(unchecked((short)0xC141), p3.Id);
            Assert.AreEqual("[01, 02, 03]", HexDump.ToHex(p3.ComplexData));


        }



    }
}