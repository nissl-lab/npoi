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
using System.Reflection;

using NPOI.Util;
using System;
using NUnit.Framework;
namespace TestCases.HWPF.Model
{

    [TestFixture]
    public class TestDocumentProperties
    {
        private DocumentProperties _documentProperties = null;
        private HWPFDocFixture _hWPFDocFixture;

        [Test]
        public void TestReadWrite()
        {
            int size = _documentProperties.GetSize();
            byte[] buf = new byte[size];

            _documentProperties.Serialize(buf, 0);

            DocumentProperties newDocProperties =
              new DocumentProperties(buf, 0);

            FieldInfo[] fields = typeof(DocumentProperties).BaseType.GetFields(BindingFlags.Instance | BindingFlags.NonPublic);

            for (int x = 0; x < fields.Length; x++)
            {
                if (!fields[x].FieldType.IsArray)
                {
                    Assert.AreEqual(fields[x].GetValue(_documentProperties),
                                 fields[x].GetValue(newDocProperties));
                }
                else
                {
                    byte[] buf1 = (byte[])fields[x].GetValue(_documentProperties);
                    byte[] buf2 = (byte[])fields[x].GetValue(newDocProperties);
                    Arrays.Equals(buf1, buf2);
                }
            }

        }
        [SetUp]
        public void SetUp()
        {
            _hWPFDocFixture = new HWPFDocFixture(this);

            _hWPFDocFixture.SetUp();

            _documentProperties = new DocumentProperties(_hWPFDocFixture._tableStream, _hWPFDocFixture._fib.GetFcDop());
        }
        [TearDown]
        public void TearDown()
        {
            _documentProperties = null;

            _hWPFDocFixture = null;
        }

    }
}
