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
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestEscherProperty
    {
        /**
         * assure that EscherProperty.getName() returns correct name for complex properties
         * See Bugzilla 50401 
         */
        [Test]
        public void TestPropertyNames()
        {
            EscherProperty p1 = new EscherSimpleProperty(EscherProperties.GROUPSHAPE__SHAPENAME, 0);
            Assert.AreEqual("groupshape.shapename", p1.Name);
            Assert.AreEqual(EscherProperties.GROUPSHAPE__SHAPENAME, p1.PropertyNumber);
            Assert.IsFalse(p1.IsComplex);

            EscherProperty p2 = new EscherComplexProperty(
                    EscherProperties.GROUPSHAPE__SHAPENAME, false, new byte[10]);
            Assert.AreEqual("groupshape.shapename", p2.Name);
            Assert.AreEqual(EscherProperties.GROUPSHAPE__SHAPENAME, p2.PropertyNumber);
            Assert.IsTrue(p2.IsComplex);
            Assert.IsFalse(p2.IsBlipId);

            EscherProperty p3 = new EscherComplexProperty(
                    EscherProperties.GROUPSHAPE__SHAPENAME, true, new byte[10]);
            Assert.AreEqual("groupshape.shapename", p3.Name);
            Assert.AreEqual(EscherProperties.GROUPSHAPE__SHAPENAME, p3.PropertyNumber);
            Assert.IsTrue(p3.IsComplex);
            Assert.IsTrue(p3.IsBlipId);
        }
    }
}