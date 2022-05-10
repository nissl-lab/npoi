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

using NPOI.OpenXml4Net.OPC;
using System;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.Util;
using NPOI.HSSF;
using TestCases.HSSF;
using System.Diagnostics;
using NUnit.Framework;

namespace NPOI.XSSF
{

    /**
     * Centralises logic for Finding/opening sample files in the src/testcases/org/apache/poi/hssf/hssf/data folder. 
     * 
     * @author Josh Micich
     */
    public class XSSFTestDataSamples
    {
        /**
         * Used by {@link writeOutAndReadBack(R wb, String testName)}.  If a
         * value is set for this in the System Properties, the xlsx file
         * will be written out to that directory.
         */
        public static String TEST_OUTPUT_DIR = "poi.test.xssf.output.dir";

        public static FileInfo GetSampleFile(String sampleFileName)
        {
            return HSSFTestDataSamples.GetSampleFile(sampleFileName);
        }
        public static OPCPackage OpenSamplePackage(String sampleName)
        {
            return OPCPackage.Open(
                  HSSFTestDataSamples.OpenSampleFileStream(sampleName)
            );
        }
        public static XSSFWorkbook OpenSampleWorkbook(String sampleName)
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(sampleName);
            return new XSSFWorkbook(is1);

        }
        public static IWorkbook WriteOutAndReadBack(IWorkbook wb)
        {
            IWorkbook result;
            try
            {
                using (MemoryStream baos = new MemoryStream(8192))
                {
                    Stopwatch sw = new Stopwatch();
                    sw.Start();
                    wb.Write(baos, false);
                    sw.Stop();
                    Debug.WriteLine("XSSFWorkbook write time: " + sw.ElapsedMilliseconds + "ms");
                    using (Stream is1 = new MemoryStream(baos.ToArray()))
                    {
                        if (wb is HSSFWorkbook)
                        {
                            result = new HSSFWorkbook(is1);
                        }
                        else if (wb is XSSFWorkbook)
                        {
                            Stopwatch sw2 = new Stopwatch();
                            sw2.Start();
                            result = new XSSFWorkbook(is1);
                            sw2.Stop();
                            Debug.WriteLine("XSSFWorkbook parse time: " + sw2.ElapsedMilliseconds + "ms");
                        }
                        else
                        {
                            throw new RuntimeException("Unexpected workbook type ("
                                    + wb.GetType().Name + ")");
                        }
                    }
                }
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
            return result;
        }

        /**
         * Writes the Workbook either into a file or into a byte array, depending on presence of 
         * the system property {@value #TEST_OUTPUT_DIR}, and reads it in a new instance of the Workbook back.
         * @param wb workbook to write
         * @param testName file name to be used if writing into a file. The old file with the same name will be overridden.
         * @return new instance read from the stream written by the wb parameter.
         */
        public static XSSFWorkbook WriteOutAndReadBack(XSSFWorkbook wb, String testName)
        {
            XSSFWorkbook result = null;


            try
            {
                string filename = Path.Combine(TestContext.CurrentContext.TestDirectory, testName + ".xlsx");

                if (File.Exists(filename))
                {
                    File.Delete(filename);
                }
                FileStream out1 = new FileStream(filename, FileMode.Create, FileAccess.ReadWrite);
                try
                {
                    wb.Write(out1);
                }
                finally
                {
                    out1.Close();
                }
                FileStream in1 = new FileStream(filename, FileMode.Open, FileAccess.Read);
                try
                {
                    result = new XSSFWorkbook(in1);
                }
                finally
                {
                    in1.Close();
                }
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }

            return result;
        }


        /**
         * Write out workbook <code>wb</code> to {@link #TEST_OUTPUT_DIR}/testName.xlsx
         * (or create a temporary file if <code>TEST_OUTPUT_DIR</code> is not defined).
         *
         * @param wb the workbook to write
         * @param testName a fragment of the filename
         * @return the location where the workbook was saved
         * @throws IOException
         */
        public static FileInfo WriteOut<R>(R wb, String testName) where R : IWorkbook
        {
            String testOutputDir = TestContext.Parameters[TEST_OUTPUT_DIR];
            FileInfo file;
            if (testOutputDir != null)
            {
                file = new FileInfo(Path.Combine(testOutputDir, testName + ".xlsx"));
            }
            else
            {
                file = TempFile.CreateTempFile(testName, ".xlsx");
            }
            if (file.Exists)
            {
                file.Delete();
            }
            FileStream out1 = file.Create();
            try
            {
                wb.Write(out1, false);
            }
            finally
            {
                out1.Close();
            }
            return file;
        }

        /**
         * Write out workbook <code>wb</code> to a memory buffer
         *
         * @param wb the workbook to write
         * @return the memory buffer
         * @throws IOException
         */
        public static ByteArrayOutputStream WriteOut<R>(R wb) where R : IWorkbook
        {
            ByteArrayOutputStream out1 = new ByteArrayOutputStream(8192);
            wb.Write(out1, false);
            return out1;
        }

        /**
         * Write out the workbook then closes the workbook. 
         * This should be used when there is insufficient memory to have
         * both workbooks open.
         * 
         * Make sure there are no references to any objects in the workbook
         * so that garbage collection may free the workbook.
         * 
         * After calling this method, null the reference to <code>wb</code>,
         * then call {@link #readBack(File)} or {@link #readBackAndDelete(File)} to re-read the file.
         * 
         * Alternatively, use {@link #writeOutAndClose(Workbook)} to use a ByteArrayOutputStream/ByteArrayInputStream
         * to avoid creating a temporary file. However, this may complicate the calling
         * code to avoid having the workbook, BAOS, and BAIS open at the same time.
         *
         * @param wb
         * @param testName file name to be used to write to a file. This file will be cleaned up by a call to readBack(String)
         * @return workbook location
         * @throws RuntimeException if {@link #TEST_OUTPUT_DIR} System property is not set
         */
        public static FileInfo WriteOutAndClose<R>(R wb, String testName) where R : IWorkbook
        {
            try
            {
                FileInfo file = WriteOut(wb, testName);
                // Do not close the workbook if there was a problem writing the workbook
                wb.Close();
                return file;
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
        }


        /**
         * Write out workbook <code>wb</code> to a memory buffer,
         * then close the workbook
         *
         * @param wb the workbook to write
         * @return the memory buffer
         * @throws IOException
         */
        public static ByteArrayOutputStream WriteOutAndClose<R>(R wb) where R : IWorkbook
        {
            try
            {
                ByteArrayOutputStream out1 = WriteOut(wb);
                // Do not close the workbook if there was a problem writing the workbook
                wb.Close();
                return out1;
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
        }

        /**
         * Read back a workbook that was written out to a file with
         * {@link #writeOut(Workbook, String))} or {@link #writeOutAndClose(Workbook, String)}.
         * Deletes the file after reading back the file.
         * Does not delete the file if an exception is raised.
         *
         * @param file the workbook file to read and delete
         * @return the read back workbook
         * @throws IOException
         */
        public static XSSFWorkbook ReadBackAndDelete(FileInfo file)
        {
            XSSFWorkbook wb = ReadBack(file);
            // do not delete the file if there's an error--might be helpful for debugging 
            file.Delete();
            return wb;
        }

        /**
         * Read back a workbook that was written out to a file with
         * {@link #writeOut(Workbook, String)} or {@link #writeOutAndClose(Workbook, String)}.
         *
         * @param file the workbook file to read
         * @return the read back workbook
         * @throws IOException
         */
        public static XSSFWorkbook ReadBack(FileInfo file)
        {
            FileStream in1 = file.Create();
            try
            {
                return new XSSFWorkbook(in1);
            }
            finally
            {
                in1.Close();
            }
        }

        /**
         * Read back a workbook that was written out to a memory buffer with
         * {@link #writeOut(Workbook)} or {@link #writeOutAndClose(Workbook)}.
         *
         * @param file the workbook file to read
         * @return the read back workbook
         * @throws IOException
         */
        public static XSSFWorkbook ReadBack(ByteArrayOutputStream out1)
        {
            InputStream is1 = new ByteArrayInputStream(out1.ToByteArray());
            out1.Close();
            try
            {
                return new XSSFWorkbook(is1);
            }
            finally
            {
                is1.Close();
            }
        }

        /**
         * Write out and read back using a memory buffer to avoid disk I/O.
         * If there is not enough memory to have two workbooks open at the same time,
         * consider using:
         * 
         * Workbook wb = new XSSFWorkbook();
         * String testName = "example";
         * 
         * <code>
         * File file = writeOutAndClose(wb, testName);
         * // clear all references that would prevent the workbook from getting garbage collected
         * wb = null;
         * Workbook wbBack = readBackAndDelete(file);
         * </code>
         *
         * @param wb the workbook to write out
         * @return the read back workbook
         */
        public static R WriteOutAndReadBack<R>(R wb) where R : IWorkbook
        {
            IWorkbook result;
            try
            {
                result = ReadBack(WriteOut(wb));
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
            R r = (R)result;
            return r;
        }

        /**
         * Write out, close, and read back the workbook using a memory buffer to avoid disk I/O.
         *
         * @param wb the workbook to write out and close
         * @return the read back workbook
         */
        public static R WriteOutCloseAndReadBack<R>(R wb) where R : IWorkbook
        {
            IWorkbook result;
            try
            {
                result = ReadBack(WriteOutAndClose(wb));
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }

            R r = (R)result;
            return r;

        }

        /**
         * Writes the Workbook either into a file or into a byte array, depending on presence of 
         * the system property {@value #TEST_OUTPUT_DIR}, and reads it in a new instance of the Workbook back.
         * If TEST_OUTPUT_DIR is set, the file will NOT be deleted at the end of this function.
         * @param wb workbook to write
         * @param testName file name to be used if writing into a file. The old file with the same name will be overridden.
         * @return new instance read from the stream written by the wb parameter.
         */

        public static R WriteOutAndReadBack<R>(R wb, String testName) where R : IWorkbook
        {
            if (TestContext.Parameters[TEST_OUTPUT_DIR] == null)
            {
                return WriteOutAndReadBack(wb);
            }
            else
            {
                try
                {
                    IWorkbook result = ReadBack(WriteOut(wb, testName));
                    R r = (R)result;
                    return r;
                }
                catch (IOException e)
                {
                    throw new RuntimeException(e);
                }
            }
        }

    }

}

