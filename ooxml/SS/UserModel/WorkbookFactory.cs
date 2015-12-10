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
    public enum ImportOption
    {
        NONE,
        /// <summary>
        /// Only Text and Formulas are imported. Pictures, Drawing, Styles etc. are all ignored.
        /// </summary>
        SheetContentOnly,
        /// <summary>
        /// Only Text, Comments and Formulas are imported. Pictures, Drawing, Styles etc. are all ignored.
        /// </summary>
        TextOnly,
        /// <summary>
        /// Everything is imported - this is the same as NONE.
        /// </summary>
        All,
    }

    /// <summary>
    /// Factory for creating the appropriate kind of Workbook
    /// (be it HSSFWorkbook or XSSFWorkbook), from the given input
    /// </summary>
    public class WorkbookFactory
    {
        /// <summary>
        /// Creates an HSSFWorkbook from the given POIFSFileSystem
        /// </summary>
        public static IWorkbook Create(POIFSFileSystem fs)
        {
            return new HSSFWorkbook(fs);
        }

        /**
         * Creates an HSSFWorkbook from the given NPOIFSFileSystem
         */
        public static IWorkbook Create(NPOIFSFileSystem fs)
        {
            return new HSSFWorkbook(fs.Root, true);
        }

        /// <summary>
        /// Creates an XSSFWorkbook from the given OOXML Package
        /// </summary>
        public static IWorkbook Create(OPCPackage pkg)
        {
            return new XSSFWorkbook(pkg);
        }

        /// <summary>
        /// Creates the appropriate HSSFWorkbook / XSSFWorkbook from
        /// the given InputStream. The Stream is wraped inside a PushbackInputStream.
        /// </summary>
        /// <param name="inputStream">Input Stream of .xls or .xlsx file</param>
        /// <returns>IWorkbook depending on the input HSSFWorkbook or XSSFWorkbook is returned.</returns>
        // Your input stream MUST either support mark/reset, or
        //  be wrapped as a {@link PushbackInputStream}!
        public static IWorkbook Create(Stream inputStream)
        {
            // If Clearly doesn't do mark/reset, wrap up
            //if (!inp.MarkSupported())
            //{
            //    inp = new PushbackInputStream(inp, 8);
            //}
            inputStream = new PushbackStream(inputStream);
            if (POIFSFileSystem.HasPOIFSHeader(inputStream))
            {
                return new HSSFWorkbook(inputStream);
            }
            inputStream.Position = 0;
            if (POIXMLDocument.HasOOXMLHeader(inputStream))
            {
                return new XSSFWorkbook(OPCPackage.Open(inputStream));
            }
            throw new ArgumentException("Your stream was neither an OLE2 stream, nor an OOXML stream.");
        }
        /**
        * Creates the appropriate HSSFWorkbook / XSSFWorkbook from
        *  the given File, which must exist and be readable.
        * <p>Note that for Workbooks opened this way, it is not possible
        *  to explicitly close the underlying File resource.
        */
        public static IWorkbook Create(string file)
        {
            if (!File.Exists(file))
            {
                throw new FileNotFoundException(file);
            }
            FileStream fStream = null;
            try
            {
                using (fStream = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook wb = new HSSFWorkbook(fStream);
                    return wb;
                }
            }
            catch (OfficeXmlFileException e)
            {
                // opening as .xls failed => try opening as .xlsx
                OPCPackage pkg = OPCPackage.Open(file);
                try
                {
                    return new XSSFWorkbook(pkg);
                }
                catch (IOException ioe)
                {
                    // ensure that file handles are closed (use revert() to not re-write the file)
                    pkg.Revert();
                    //pkg.close();

                    // rethrow exception
                    throw ioe;
                }
                catch (ArgumentException ioe)
                {
                    // ensure that file handles are closed (use revert() to not re-write the file) 
                    pkg.Revert();
                    //pkg.close();

                    // rethrow exception
                    throw ioe;
                }
            }
        }
        /// <summary>
        /// Creates the appropriate HSSFWorkbook / XSSFWorkbook from
        /// the given InputStream. The Stream is wraped inside a PushbackInputStream.
        /// </summary>
        /// <param name="inputStream">Input Stream of .xls or .xlsx file</param>
        /// <param name="importOption">Customize the elements that are processed on the next import</param>
        /// <returns>IWorkbook depending on the input HSSFWorkbook or XSSFWorkbook is returned.</returns>
        public static IWorkbook Create(Stream inputStream, ImportOption importOption)
        {
            SetImportOption(importOption);
            IWorkbook workbook = Create(inputStream);
            return workbook;
        }

        /// <summary>
        /// Creates a specific FormulaEvaluator for the given workbook.
        /// </summary>
        public static IFormulaEvaluator CreateFormulaEvaluator(IWorkbook workbook)
        {
            if (typeof(HSSFWorkbook) == workbook.GetType())
            {
                return new HSSFFormulaEvaluator(workbook as HSSFWorkbook);
            }
            else
            {
                return new XSSFFormulaEvaluator(workbook as XSSFWorkbook);
            }
        }

        /// <summary>
        /// Sets the import option when opening the next workbook.
        /// Works only for XSSF. For HSSF workbooks this option is ignored.
        /// </summary>
        /// <param name="importOption">Customize the elements that are processed on the next import</param>
        public static void SetImportOption(ImportOption importOption)
        {
            if (ImportOption.SheetContentOnly == importOption)
            {
                // Add
                XSSFRelation.AddRelation(XSSFRelation.WORKSHEET);
                XSSFRelation.AddRelation(XSSFRelation.SHARED_STRINGS);

                // Remove
                XSSFRelation.RemoveRelation(XSSFRelation.WORKBOOK);
                XSSFRelation.RemoveRelation(XSSFRelation.MACROS_WORKBOOK);
                XSSFRelation.RemoveRelation(XSSFRelation.TEMPLATE_WORKBOOK);
                XSSFRelation.RemoveRelation(XSSFRelation.MACRO_TEMPLATE_WORKBOOK);
                XSSFRelation.RemoveRelation(XSSFRelation.MACRO_ADDIN_WORKBOOK);
                XSSFRelation.RemoveRelation(XSSFRelation.CHARTSHEET);
                XSSFRelation.RemoveRelation(XSSFRelation.STYLES);
                XSSFRelation.RemoveRelation(XSSFRelation.DRAWINGS);
                XSSFRelation.RemoveRelation(XSSFRelation.CHART);
                XSSFRelation.RemoveRelation(XSSFRelation.VML_DRAWINGS);
                XSSFRelation.RemoveRelation(XSSFRelation.CUSTOM_XML_MAPPINGS);
                XSSFRelation.RemoveRelation(XSSFRelation.TABLE);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGES);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGE_EMF);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGE_WMF);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGE_PICT);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGE_JPEG);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGE_PNG);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGE_DIB);
                XSSFRelation.RemoveRelation(XSSFRelation.SHEET_COMMENTS);
                XSSFRelation.RemoveRelation(XSSFRelation.SHEET_HYPERLINKS);
                XSSFRelation.RemoveRelation(XSSFRelation.OLEEMBEDDINGS);
                XSSFRelation.RemoveRelation(XSSFRelation.PACKEMBEDDINGS);
                XSSFRelation.RemoveRelation(XSSFRelation.VBA_MACROS);
                XSSFRelation.RemoveRelation(XSSFRelation.ACTIVEX_CONTROLS);
                XSSFRelation.RemoveRelation(XSSFRelation.ACTIVEX_BINS);
                XSSFRelation.RemoveRelation(XSSFRelation.THEME);
                XSSFRelation.RemoveRelation(XSSFRelation.CALC_CHAIN);
                XSSFRelation.RemoveRelation(XSSFRelation.PRINTER_SETTINGS);
            }
            else if (ImportOption.TextOnly == importOption)
            {
                // Add
                XSSFRelation.AddRelation(XSSFRelation.WORKSHEET);
                XSSFRelation.AddRelation(XSSFRelation.SHARED_STRINGS);
                XSSFRelation.AddRelation(XSSFRelation.SHEET_COMMENTS);

                // Remove
                XSSFRelation.RemoveRelation(XSSFRelation.WORKBOOK);
                XSSFRelation.RemoveRelation(XSSFRelation.MACROS_WORKBOOK);
                XSSFRelation.RemoveRelation(XSSFRelation.TEMPLATE_WORKBOOK);
                XSSFRelation.RemoveRelation(XSSFRelation.MACRO_TEMPLATE_WORKBOOK);
                XSSFRelation.RemoveRelation(XSSFRelation.MACRO_ADDIN_WORKBOOK);
                XSSFRelation.RemoveRelation(XSSFRelation.CHARTSHEET);
                XSSFRelation.RemoveRelation(XSSFRelation.STYLES);
                XSSFRelation.RemoveRelation(XSSFRelation.DRAWINGS);
                XSSFRelation.RemoveRelation(XSSFRelation.CHART);
                XSSFRelation.RemoveRelation(XSSFRelation.VML_DRAWINGS);
                XSSFRelation.RemoveRelation(XSSFRelation.CUSTOM_XML_MAPPINGS);
                XSSFRelation.RemoveRelation(XSSFRelation.TABLE);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGES);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGE_EMF);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGE_WMF);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGE_PICT);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGE_JPEG);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGE_PNG);
                XSSFRelation.RemoveRelation(XSSFRelation.IMAGE_DIB);
                XSSFRelation.RemoveRelation(XSSFRelation.SHEET_HYPERLINKS);
                XSSFRelation.RemoveRelation(XSSFRelation.OLEEMBEDDINGS);
                XSSFRelation.RemoveRelation(XSSFRelation.PACKEMBEDDINGS);
                XSSFRelation.RemoveRelation(XSSFRelation.VBA_MACROS);
                XSSFRelation.RemoveRelation(XSSFRelation.ACTIVEX_CONTROLS);
                XSSFRelation.RemoveRelation(XSSFRelation.ACTIVEX_BINS);
                XSSFRelation.RemoveRelation(XSSFRelation.THEME);
                XSSFRelation.RemoveRelation(XSSFRelation.CALC_CHAIN);
                XSSFRelation.RemoveRelation(XSSFRelation.PRINTER_SETTINGS);
            }
            else
            {
                // NONE/All
                XSSFRelation.AddRelation(XSSFRelation.WORKBOOK);
                XSSFRelation.AddRelation(XSSFRelation.MACROS_WORKBOOK);
                XSSFRelation.AddRelation(XSSFRelation.TEMPLATE_WORKBOOK);
                XSSFRelation.AddRelation(XSSFRelation.MACRO_TEMPLATE_WORKBOOK);
                XSSFRelation.AddRelation(XSSFRelation.MACRO_ADDIN_WORKBOOK);
                XSSFRelation.AddRelation(XSSFRelation.WORKSHEET);
                XSSFRelation.AddRelation(XSSFRelation.CHARTSHEET);
                XSSFRelation.AddRelation(XSSFRelation.SHARED_STRINGS);
                XSSFRelation.AddRelation(XSSFRelation.STYLES);
                XSSFRelation.AddRelation(XSSFRelation.DRAWINGS);
                XSSFRelation.AddRelation(XSSFRelation.CHART);
                XSSFRelation.AddRelation(XSSFRelation.VML_DRAWINGS);
                XSSFRelation.AddRelation(XSSFRelation.CUSTOM_XML_MAPPINGS);
                XSSFRelation.AddRelation(XSSFRelation.TABLE);
                XSSFRelation.AddRelation(XSSFRelation.IMAGES);
                XSSFRelation.AddRelation(XSSFRelation.IMAGE_EMF);
                XSSFRelation.AddRelation(XSSFRelation.IMAGE_WMF);
                XSSFRelation.AddRelation(XSSFRelation.IMAGE_PICT);
                XSSFRelation.AddRelation(XSSFRelation.IMAGE_JPEG);
                XSSFRelation.AddRelation(XSSFRelation.IMAGE_PNG);
                XSSFRelation.AddRelation(XSSFRelation.IMAGE_DIB);
                XSSFRelation.AddRelation(XSSFRelation.SHEET_COMMENTS);
                XSSFRelation.AddRelation(XSSFRelation.SHEET_HYPERLINKS);
                XSSFRelation.AddRelation(XSSFRelation.OLEEMBEDDINGS);
                XSSFRelation.AddRelation(XSSFRelation.PACKEMBEDDINGS);
                XSSFRelation.AddRelation(XSSFRelation.VBA_MACROS);
                XSSFRelation.AddRelation(XSSFRelation.ACTIVEX_CONTROLS);
                XSSFRelation.AddRelation(XSSFRelation.ACTIVEX_BINS);
                XSSFRelation.AddRelation(XSSFRelation.THEME);
                XSSFRelation.AddRelation(XSSFRelation.CALC_CHAIN);
                XSSFRelation.AddRelation(XSSFRelation.PRINTER_SETTINGS);
            }
        }

    }

}