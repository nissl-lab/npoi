using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NPOI.HPSF;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestCases.HSSF;

namespace TestCases
{
    [TestFixture]
    public class TestPOIDocumentMain
    {
        // The POI Documents to work on
        private POIDocument doc;
        private POIDocument doc2;

        /**
         * Set things up, two spreadsheets for our testing
         */
        [SetUp]
        public void setUp()
        {
            doc = HSSFTestDataSamples.OpenSampleWorkbook("DateFormats.xls");
            doc2 = HSSFTestDataSamples.OpenSampleWorkbook("StringFormulas.xls");
        }
        [Test]
        public void TestReadProperties()
        {
            readPropertiesHelper(doc);
        }

        private void readPropertiesHelper(POIDocument docWB)
        {
            // We should have both sets
            Assert.IsNotNull(docWB.DocumentSummaryInformation);
            Assert.IsNotNull(docWB.SummaryInformation);

            // Check they are as expected for the test doc
            Assert.AreEqual("Administrator", docWB.SummaryInformation.Author);
            Assert.AreEqual(0, docWB.DocumentSummaryInformation.ByteCount);
        }
        [Test]
        public void TestReadProperties2()
        {
            // Check again on the word one
            Assert.IsNotNull(doc2.DocumentSummaryInformation);
            Assert.IsNotNull(doc2.SummaryInformation);

            Assert.AreEqual("Avik Sengupta", doc2.SummaryInformation.Author);
            Assert.AreEqual(null, doc2.SummaryInformation.Keywords);
            Assert.AreEqual(0, doc2.DocumentSummaryInformation.ByteCount);
        }
        [Test]
        public void TestWriteProperties()
        {
            // Just check we can write them back out into a filesystem
            NPOIFSFileSystem outFS = new NPOIFSFileSystem();
            doc.ReadProperties();
            doc.WriteProperties(outFS);

            // Should now hold them
            Assert.IsNotNull(
                    outFS.CreateDocumentInputStream("\x0005SummaryInformation")
            );
            Assert.IsNotNull(
                    outFS.CreateDocumentInputStream("\x0005DocumentSummaryInformation")
            );
        }
        [Test]
        public void TestWriteReadProperties()
        {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            // Write them out
            NPOIFSFileSystem outFS = new NPOIFSFileSystem();
            doc.ReadProperties();
            doc.WriteProperties(outFS);
            outFS.WriteFileSystem(baos);

            // Create a new version
            ByteArrayInputStream bais = new ByteArrayInputStream(baos.ToByteArray());
            OPOIFSFileSystem inFS = new OPOIFSFileSystem(bais);

            // Check they're still there
            POIDocument doc3 = new HPSFPropertiesOnlyDocument(inFS);
            doc3.ReadProperties();

            // Delegate test
            readPropertiesHelper(doc3);
            doc3.Close();
        }
        [Test]
        public void TestCreateNewProperties()        {
            POIDocument doc = new HSSFWorkbook();

            // New document won't have them
            Assert.IsNull(doc.SummaryInformation);
            Assert.IsNull(doc.DocumentSummaryInformation);

            // Add them in
            doc.CreateInformationProperties();
            Assert.IsNotNull(doc.SummaryInformation);
            Assert.IsNotNull(doc.DocumentSummaryInformation);

            // Write out and back in again, no change
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            doc.Write(baos);
            ByteArrayInputStream bais = new ByteArrayInputStream(baos.ToByteArray());

            doc = new HSSFWorkbook(bais);

            Assert.IsNotNull(doc.SummaryInformation);
            Assert.IsNotNull(doc.DocumentSummaryInformation);
        }
        [Test]
        public void TestCreateNewPropertiesOnExistingFile()
        {
            POIDocument doc = new HSSFWorkbook();

            // New document won't have them
            Assert.IsNull(doc.SummaryInformation);
            Assert.IsNull(doc.DocumentSummaryInformation);

            // Write out and back in again, no change
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            doc.Write(baos);
            ByteArrayInputStream bais = new ByteArrayInputStream(baos.ToByteArray());
            doc = new HSSFWorkbook(bais);

            Assert.IsNull(doc.SummaryInformation);
            Assert.IsNull(doc.DocumentSummaryInformation);

            // Create, and change
            doc.CreateInformationProperties();
            doc.SummaryInformation.Author = ("POI Testing");
            doc.DocumentSummaryInformation.Company = ("ASF");

            // Save and re-load
            baos = new ByteArrayOutputStream();
            doc.Write(baos);
            bais = new ByteArrayInputStream(baos.ToByteArray());
            doc = new HSSFWorkbook(bais);

            // Check
            Assert.IsNotNull(doc.SummaryInformation);
            Assert.IsNotNull(doc.DocumentSummaryInformation);
            Assert.AreEqual("POI Testing", doc.SummaryInformation.Author);
            Assert.AreEqual("ASF", doc.DocumentSummaryInformation.Company);

            // Asking to re-create will make no difference now
            doc.CreateInformationProperties();
            Assert.IsNotNull(doc.SummaryInformation);
            Assert.IsNotNull(doc.DocumentSummaryInformation);
            Assert.AreEqual("POI Testing", doc.SummaryInformation.Author);
            Assert.AreEqual("ASF", doc.DocumentSummaryInformation.Company);
        }
    }
}
