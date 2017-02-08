using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.XSSF.Streaming;
using NUnit.Framework;

namespace NPOI.OOXML.Testcases.XSSF.Streaming
{
    [TestFixture]
    class SheetDataWriterTests
    {
        private SheetDataWriter _objectToTest;

        [Test]
        public void IfCallingEmptyConstructorShouldCreateNonZippedTempFileNonDecoratedStream()
        {
            _objectToTest = new SheetDataWriter();
            Assert.True(_objectToTest.TemporaryFilePath().Contains("poi-sxssf-sheet"));
            Assert.True(!_objectToTest.TemporaryFilePath().Contains(".gz"));
            if (File.Exists(_objectToTest.TemporaryFilePath()))
            {
                _objectToTest.Dispose();
                File.Delete(_objectToTest.TemporaryFilePath());
            }
        }


    }
}
