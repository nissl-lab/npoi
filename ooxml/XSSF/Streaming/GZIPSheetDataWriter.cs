using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ICSharpCode.SharpZipLib.GZip;
using NPOI.Util;
using NPOI.XSSF.Model;

namespace NPOI.XSSF.Streaming
{
    public class GZIPSheetDataWriter : SheetDataWriter
    {
        public GZIPSheetDataWriter() :base()
        {
           
        }

        /**
         * @param sharedStringsTable the shared strings table, or null if inline text is used
         */
        public GZIPSheetDataWriter(SharedStringsTable sharedStringsTable) : base(sharedStringsTable)
        {
            
        }

        /**
         * @return temp file to write sheet data
         */

    public FileInfo createTempFile()
        {
        return TempFile.CreateTempFile("poi-sxssf-sheet-xml", ".gz");
        }


    protected Stream decorateInputStream(Stream fis)
        {
        return new GZipInputStream(fis);
    }

    protected Stream decorateOutputStream(Stream fos)
    {
        return new GZipOutputStream(fos);
}
    }
}
