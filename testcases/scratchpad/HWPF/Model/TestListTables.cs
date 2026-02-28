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

using NPOI.HWPF.Model;
using NPOI.HWPF.Model.IO;
using NUnit.Framework;

namespace TestCases.HWPF.Model
{
    [TestFixture]
    public class TestListTables : HWPFTestCase
    {

        public TestListTables()
        {
        }
        [Test]
        public void TestReadWrite()
        {
            FileInformationBlock fib = _hWPFDocFixture._fib;
            byte[] tableStream = _hWPFDocFixture._tableStream;

            int listOffset = fib.GetFcPlcfLst();
            int lfoOffset = fib.GetFcPlfLfo();
            if (listOffset != 0 && fib.GetLcbPlcfLst() != 0)
            {
                ListTables listTables = new ListTables(tableStream, fib.GetFcPlcfLst(),
                                                        fib.GetFcPlfLfo());
                HWPFFileSystem fileSys = new HWPFFileSystem();

                HWPFStream tableOut = fileSys.GetStream("1Table");

                listTables.WriteListDataTo(tableOut);
                int offset = tableOut.Offset;
                listTables.WriteListOverridesTo(tableOut);

                ListTables newTables = new ListTables(tableOut.ToArray(), 0, offset);

                Assert.AreEqual(listTables, newTables);

            }
        }

    }
}
