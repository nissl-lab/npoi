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
using NPOI.HSSF.Record.Crypto;
using NPOI.HSSF.UserModel;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.OpenXml4Net.OPC;
using NPOI.POIFS.Crypt;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Security;

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
    public static class WorkbookFactory
    {
        /// <summary>
        /// Creates an HSSFWorkbook from the given POIFSFileSystem
        /// </summary>
        public static IWorkbook Create(POIFSFileSystem fs)
        {
            return new HSSFWorkbook(fs);
        }

        /// <summary>
        /// Creates an HSSFWorkbook from the given NPOIFSFileSystem
        /// </summary>
        /// <param name="fs"></param>
        /// <returns></returns>
        public static IWorkbook Create(NPOIFSFileSystem fs)
        {
            return new HSSFWorkbook(fs.Root, true);
        }

        /// <summary>
        /// Creates a Workbook from the given NPOIFSFileSystem, which may be password protected
        /// </summary>
        /// <param name="fs"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        private static IWorkbook Create(NPOIFSFileSystem fs, string password)
        {
            DirectoryNode root = fs.Root;

            // Encrypted OOXML files go inside OLE2 containers, is this one?
            if (root.HasEntry(Decryptor.DEFAULT_POIFS_ENTRY))
            {
                InputStream stream = DocumentFactoryHelper.GetDecryptedStream(fs, password);

                OPCPackage pkg = OPCPackage.Open(stream);
                return Create(pkg);
            }

            // If we get here, it isn't an encrypted XLSX file
            // So, treat it as a regular HSSF XLS one
            if (password != null)
            {
                Biff8EncryptionKey.CurrentUserPassword = (password);
            }
            try
            {
                return new HSSFWorkbook(root, true);
            }
            finally
            {
                Biff8EncryptionKey.CurrentUserPassword = (null);
            }
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
        /// <remarks>Your input stream MUST either support mark/reset, or
        ///  be wrapped as a {@link PushbackInputStream}!</remarks>
        public static IWorkbook Create(Stream inputStream, bool bReadonly)
        {
            if (inputStream.Length == 0)
                throw new EmptyFileException();
            inputStream = new PushbackStream(inputStream);
            if (POIFSFileSystem.HasPOIFSHeader(inputStream))
            {
                return new HSSFWorkbook(inputStream);
            }
            inputStream.Position = 0;
            if (DocumentFactoryHelper.HasOOXMLHeader(inputStream))
            {
                return new XSSFWorkbook(OPCPackage.Open(inputStream, bReadonly));
            }
            throw new InvalidFormatException("Your stream was neither an OLE2 stream, nor an OOXML stream.");
        }

        public static IWorkbook Create(Stream inputStream)
        {
            return Create(inputStream, false);
        }

        public static IWorkbook Create(string file, string password, bool readOnly)
        {
            if (!File.Exists(file))
            {
                throw new FileNotFoundException(file);
            }

            Stream inputStream = new FileStream(file, FileMode.Open, readOnly ? FileAccess.Read : FileAccess.ReadWrite);

            try
            {
                return Create(inputStream, password, readOnly);
            }
            finally
            {
                if (inputStream != null)
                    inputStream.Dispose();
            }
        }

        public static IWorkbook Create(string file)
        {
            if (!File.Exists(file))
            {
                throw new FileNotFoundException(file);
            }
            using (var stream = File.OpenRead(file))
            {
                return Create(stream);
            }
        }

        public static IWorkbook Create(string file, string password)
        {
            return Create(file, password, false);
        }

        /// <summary>
        /// Creates the appropriate HSSFWorkbook / XSSFWorkbook from 
        /// the given File, which must exist and be readable.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="password"></param>
        /// <param name="readOnly"></param>
        /// <returns></returns>
        /// <remarks>
        /// Note that for Workbooks opened this way, it is not possible
        /// to explicitly close the underlying File resource.
        /// </remarks>
        public static IWorkbook Create(Stream inputStream, string password, bool readOnly)
        {
            inputStream = new PushbackStream(inputStream);

            try
            {
                if (POIFSFileSystem.HasPOIFSHeader(inputStream))
                {
                    if (DocumentFactoryHelper.GetPasswordProtected(inputStream) == DocumentFactoryHelper.OfficeProtectType.ProtectedOOXML)
                    {
                        inputStream.Position = 0;
                        POIFSFileSystem fs = new POIFSFileSystem(inputStream);

                        var decriptedStream = DocumentFactoryHelper.GetDecryptedStream(fs, password);

                        return new XSSFWorkbook(decriptedStream);
                    }

                    inputStream.Position = 0;
                    NPOIFSFileSystem nfs = new NPOIFSFileSystem(inputStream);

                    try
                    {
                        return Create(nfs, password);
                    }
                    finally
                    {
                        // ensure that the file-handle is closed again
                        if (nfs != null)
                            nfs.Close();
                    }
                }

                inputStream.Position = 0;
                if (DocumentFactoryHelper.HasOOXMLHeader(inputStream))
                {
                    inputStream.Position = 0;
                    OPCPackage pkg = OPCPackage.Open(inputStream, readOnly);
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
                    catch (Exception ioe)
                    {
                        // ensure that file handles are closed (use revert() to not re-write the file) 
                        pkg.Revert();
                        //pkg.close();

                        // rethrow exception
                        throw ioe;
                    }
                }
            }
            finally
            {
                if (inputStream != null)
                    inputStream.Dispose();
            }

            throw new InvalidFormatException("Your stream was neither an OLE2 stream, nor an OOXML stream.");
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
            IWorkbook workbook = Create(inputStream, true);
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