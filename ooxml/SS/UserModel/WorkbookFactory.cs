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

using System;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.OpenXml4Net.OPC;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.SS.UserModel
{

    /**
     * Factory for creating the appropriate kind of Workbook
     *  (be it HSSFWorkbook or XSSFWorkbook), from the given input
     */
    public class WorkbookFactory
    {
        /**
         * Creates an HSSFWorkbook from the given POIFSFileSystem
         */
        public static IWorkbook Create(POIFSFileSystem fs)
        {
            return new HSSFWorkbook(fs);
        }
        /**
         * Creates an XSSFWorkbook from the given OOXML Package
         */
        public static IWorkbook Create(OPCPackage pkg)
        {
            return new XSSFWorkbook(pkg);
        }
        /**
         * Creates the appropriate HSSFWorkbook / XSSFWorkbook from
         * the given InputStream. The Stream is wraped inside
         * a PushbackInputStream.
         */
        // Your input stream MUST either support mark/reset, or
         //  be wrapped as a {@link PushbackInputStream}!
        public static IWorkbook Create(Stream inp)
        {
            // If Clearly doesn't do mark/reset, wrap up
            //if (!inp.MarkSupported())
            //{
            //    inp = new PushbackInputStream(inp, 8);
            //}
            inp = new PushbackStream(inp);
            if (POIFSFileSystem.HasPOIFSHeader(inp))
            {
                return new HSSFWorkbook(inp);
            }
            inp.Position = 0;
            if (POIXMLDocument.HasOOXMLHeader(inp))
            {
                return new XSSFWorkbook(OPCPackage.Open(inp));
            }
            throw new ArgumentException("Your InputStream was neither an OLE2 stream, nor an OOXML stream.");
        }
    }

}