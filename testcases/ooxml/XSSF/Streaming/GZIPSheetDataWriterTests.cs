using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ICSharpCode.SharpZipLib.GZip;
using NPOI.XSSF.Model;
using NPOI.XSSF.Streaming;
using NUnit.Framework;

namespace NPOI.OOXML.Testcases.XSSF.Streaming
{
    [TestFixture]
    class GZIPSheetDataWriterTests
    {
        private GZIPSheetDataWriter _objectToTest;

        [TearDown]
        public void CleanUp()
        {
            if (_objectToTest != null)
            {
                if (File.Exists(_objectToTest.TemporaryFilePath()))
                {
                    _objectToTest.Dispose();
                    File.Delete(_objectToTest.TemporaryFilePath());
                }
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
