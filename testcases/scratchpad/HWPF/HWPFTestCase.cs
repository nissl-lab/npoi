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

using System.IO;
using TestCases.HWPF;

using NPOI.HWPF;
using NUnit.Framework;
namespace TestCases.HWPF
{

    public abstract class HWPFTestCase
    {
        protected HWPFDocFixture _hWPFDocFixture;

        protected HWPFTestCase()
        {
        }
        [SetUp]
        public void SetUp()
        {

            /** @todo verify the constructors */
            _hWPFDocFixture = new HWPFDocFixture(this);

            _hWPFDocFixture.SetUp();
        }
        [TearDown]
        public void TearDown()
        {

            _hWPFDocFixture = null;
        }

        public HWPFDocument WriteOutAndRead(HWPFDocument doc)
        {
            MemoryStream baos = new MemoryStream();
            HWPFDocument newDoc;
            doc.Write(baos);
            MemoryStream bais = new MemoryStream(baos.ToArray());
            newDoc = new HWPFDocument(bais);

            return newDoc;
        }
    }

}