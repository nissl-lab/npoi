using NPOI.OpenXml4Net.OPC;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;

namespace UpdateEmbeddedDoc
{
    public class UpdateEmbeddedDoc
    {
        // this example has some bugs , office word cannot open the doc file that writed by this example.
        // TODO fix it
        private XWPFDocument doc = null;
        private string docFile = null;

        private static int SHEET_NUM = 0;
        private static int ROW_NUM = 0;
        private static int CELL_NUM = 0;
        private static double NEW_VALUE = 100.98D;
        private static String BINARY_EXTENSION = "xls";
        private static String OPENXML_EXTENSION = "xlsx";

        /**
     * Create a new instance of the UpdateEmbeddedDoc class using the following
     * parameters;
     *
     * @param filename An instance of the String class that encapsulates the name
     *                 of and path to a WordprocessingML Word document that contains an
     *                 embedded binary Excel workbook.
     * @throws java.io.FileNotFoundException Thrown if the file cannot be found
     *                                       on the underlying file system.
     * @throws java.io.IOException           Thrown if a problem occurs in the underlying
     *                                       file system.
     */
        public UpdateEmbeddedDoc(String filename)
        {
            this.docFile = filename;
            FileStream fis = null;
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException("The Word dcoument " +
                        filename +
                        " does not exist.");
            }
            try
            {

                // Open the Word document file and instantiate the XWPFDocument
                // class.
                fis = new FileStream(this.docFile, FileMode.Open);
                this.doc = new XWPFDocument(fis);
            }
            finally
            {
                if (fis != null)
                {
                    try
                    {
                        fis.Close();
                        fis = null;
                    }
                    catch (IOException)
                    {
                        Console.WriteLine("IOException caught trying to close " +
                                "FileInputStream in the constructor of " +
                                "UpdateEmbeddedDoc.");
                    }
                }
            }
        }

        /**
         * Called to update the embedded Excel workbook. As the format and structire
         * of the workbook are known in advance, all this code attempts to do is
         * write a new value into the first cell on the first row of the first
         * worksheet. Prior to executing this method, that cell will contain the
         * value 1.
         *
         * @throws org.apache.poi.openxml4j.exceptions.OpenXML4JException
         *                             Rather
         *                             than use the specific classes (HSSF/XSSF) to handle the embedded
         *                             workbook this method uses those defeined in the SS stream. As
         *                             a result, it might be the case that a SpreadsheetML file is
         *                             opened for processing, throwing this exception if that file is
         *                             invalid.
         * @throws java.io.IOException Thrown if a problem occurs in the underlying
         *                             file system.
         */
        public void UpdateEmbeddedDoc1()
        {
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            NPOI.SS.UserModel.ICell cell = null;
            PackagePart pPart = null;
            IEnumerator<PackagePart> pIter = null;
            List<PackagePart> embeddedDocs = this.doc.GetAllEmbedds();
            if (embeddedDocs != null && embeddedDocs.Count != 0)
            {
                pIter = embeddedDocs.GetEnumerator();
                while (pIter.MoveNext())
                {
                    pPart = pIter.Current;
                    if (pPart.PartName.Extension.Equals(BINARY_EXTENSION) ||
                            pPart.PartName.Extension.Equals(OPENXML_EXTENSION))
                    {

                        // Get an InputStream from the pacage part and pass that
                        // to the create method of the WorkbookFactory class. Update
                        // the resulting Workbook and then stream that out again
                        // using an OutputStream obtained from the same PackagePart.
                        workbook = WorkbookFactory.Create(pPart.GetInputStream());
                        sheet = workbook.GetSheetAt(SHEET_NUM);
                        row = sheet.GetRow(ROW_NUM);
                        cell = row.GetCell(CELL_NUM);
                        cell.SetCellValue(NEW_VALUE);
                        workbook.Write(pPart.GetOutputStream());
                    }
                }

                // Finally, write the newly modified Word document out to file.
                string file = Path.GetFileNameWithoutExtension(this.docFile) + "tmp" + Path.GetExtension(this.docFile);
                this.doc.Write(new FileStream(file, FileMode.CreateNew));
            }
        }

        /**
         * Called to test whether or not the embedded workbook was correctly
         * updated. This method simply recovers the first cell from the first row
         * of the first workbook and tests the value it contains.
         * <p/>
         * Note that execution will not continue up to the assertion as the
         * embedded workbook is now corrupted and causes an IllegalArgumentException
         * with the following message
         * <p/>
         * <em>java.lang.IllegalArgumentException: Your InputStream was neither an
         * OLE2 stream, nor an OOXML stream</em>
         * <p/>
         * to be thrown when the WorkbookFactory.createWorkbook(InputStream) method
         * is executed.
         *
         * @throws org.apache.poi.openxml4j.exceptions.OpenXML4JException
         *                             Rather
         *                             than use the specific classes (HSSF/XSSF) to handle the embedded
         *                             workbook this method uses those defeined in the SS stream. As
         *                             a result, it might be the case that a SpreadsheetML file is
         *                             opened for processing, throwing this exception if that file is
         *                             invalid.
         * @throws java.io.IOException Thrown if a problem occurs in the underlying
         *                             file system.
         */
        public void CheckUpdatedDoc()
        {
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            NPOI.SS.UserModel.ICell cell = null;
            PackagePart pPart = null;
            IEnumerator<PackagePart> pIter = null;
            List<PackagePart> embeddedDocs = this.doc.GetAllEmbedds();
            if (embeddedDocs != null && embeddedDocs.Count != 0)
            {
                pIter = embeddedDocs.GetEnumerator();
                while (pIter.MoveNext())
                {
                    pPart = pIter.Current;
                    if (pPart.PartName.Extension.Equals(BINARY_EXTENSION) ||
                            pPart.PartName.Extension.Equals(OPENXML_EXTENSION))
                    {
                        workbook = WorkbookFactory.Create(pPart.GetInputStream());
                        sheet = workbook.GetSheetAt(SHEET_NUM);
                        row = sheet.GetRow(ROW_NUM);
                        cell = row.GetCell(CELL_NUM);
                        Assert.AreEqual(cell.NumericCellValue, NEW_VALUE);
                    }
                }
            }
        }

        static void Main(string[] args)
        {
            UpdateEmbeddedDoc ued = new UpdateEmbeddedDoc(args[0]);
            ued.UpdateEmbeddedDoc1();
            ued.CheckUpdatedDoc();
        }
    }
}
