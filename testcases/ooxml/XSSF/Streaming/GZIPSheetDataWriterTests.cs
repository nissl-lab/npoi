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
using ICSharpCode.SharpZipLib.GZip;
using NPOI.XSSF.Streaming;
using NUnit.Framework;
using System.IO;

namespace TestCases.XSSF.Streaming
{
    [TestFixture]
    public class GZIPSheetDataWriterTests
    {
        private GZIPSheetDataWriter _objectToTest;

        [TearDown]
        public void CleanUp()
        {
            if (_objectToTest != null)
            {
                _objectToTest.Dispose();

                if (File.Exists(_objectToTest.TemporaryFilePath()))
                    File.Delete(_objectToTest.TemporaryFilePath());
            }
        }
        [Test]
        public void IfCallingEmptyConstructorShouldCreateZippedTempFileAndGZipOutputStream()
        {
            _objectToTest = new GZIPSheetDataWriter();
            Assert.True(_objectToTest.TemporaryFilePath().Contains("poi-sxssf-sheet-xml"));
            Assert.True(_objectToTest.TemporaryFilePath().Contains(".gz"));
        }

        [Test]
        public void IfCreatingWriterShouldCreateGZipOutPutStream()
        {
            _objectToTest = new GZIPSheetDataWriter();

            var tempFile = _objectToTest.CreateTempFile();
            using (var result = _objectToTest.CreateWriter(tempFile))
            {
                Assert.True(result is GZipOutputStream);
            }
        }

        [Test]
        public void IfGettingWorksheetXmlInputStreamShouldReturnGZipInputStream()
        {
            _objectToTest = new GZIPSheetDataWriter();
            _objectToTest.Close();

            using (var result = _objectToTest.GetWorksheetXmlInputStream())
            {
                Assert.True(result is GZipInputStream);
            }   
        }
    }
}
