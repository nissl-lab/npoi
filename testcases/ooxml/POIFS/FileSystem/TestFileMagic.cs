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

using NPOI.POIFS.FileSystem;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCases.POIFS.FileSystem
{
    [TestFixture]
    public class TestFileMagic
    {
        [Test]
        public void TestXSSFWorkbookStreamPositionGreaterThanZero()
        {
            var workbook = new XSSFWorkbook();
            workbook.CreateSheet("Nothing here");
            var stream = new MemoryStream();
            workbook.Write(stream, leaveOpen: true);
            ClassicAssert.AreEqual(FileMagic.OOXML, FileMagicContainer.ValueOf(stream));
            stream.Position = 1;
            ClassicAssert.AreEqual(FileMagic.OOXML, FileMagicContainer.ValueOf(stream));
            workbook.Close();
        }
    }
}
