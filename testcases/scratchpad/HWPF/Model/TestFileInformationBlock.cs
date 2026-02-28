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


namespace TestCases.HWPF.Model
{

    using NPOI.HWPF.Model;
    using NUnit.Framework;
    using System.Reflection;
    [TestFixture]
    public class TestFileInformationBlock
    {
        private FileInformationBlock _fileInformationBlock = null;
        private HWPFDocFixture _hWPFDocFixture;

        [Test]
        public void TestReadWrite()
        {
            int size = _fileInformationBlock.GetSize();
            byte[] buf = new byte[size];

            _fileInformationBlock.Serialize(buf, 0);

            FileInformationBlock newFileInformationBlock =
              new FileInformationBlock(buf);

            FieldInfo[] fields = typeof(FileInformationBlock).BaseType.GetFields(BindingFlags.Instance | BindingFlags.NonPublic);

            for (int x = 0; x < fields.Length; x++)
            {
                Assert.AreEqual(fields[x].GetValue(_fileInformationBlock), fields[x].GetValue(newFileInformationBlock));
            }
        }
        [SetUp]
        public void SetUp()
        {
            /**@todo verify the constructors*/
            _hWPFDocFixture = new HWPFDocFixture(this);

            _hWPFDocFixture.SetUp();
            _fileInformationBlock = _hWPFDocFixture._fib;
        }
        [TearDown]
        public void TearDown()
        {
            _fileInformationBlock = null;

            _hWPFDocFixture = null;
        }

    }

}