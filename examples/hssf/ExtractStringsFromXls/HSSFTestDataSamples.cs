/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF
{
    using System;
    using System.IO;
    using System.Text;
    using NPOI.HSSF.UserModel;

    /**
     * Centralises logic for finding/Opening sample files in the src/testcases/org/apache/poi/hssf/hssf/data folder. 
     * 
     * @author Josh Micich
     */
    public class HSSFTestDataSamples
    {

        private static String TEST_DATA_DIR_SYS_PROPERTY_NAME = "HSSF.testdata.path";

        private static string _resolvedDataDir;
        /** <code>true</code> if standard system propery is1 not set, 
         * but the data is1 available on the test runtime classpath */
        private static bool _sampleDataIsAvaliableOnClassPath;

        /**
         * Opens a sample file from the standard HSSF test data directory
         * 
         * @return an Open <tt>Stream</tt> for the specified sample file
         */
        public static Stream OpenSampleFileStream(String sampleFileName)
        {
            Initialise();

            if (_sampleDataIsAvaliableOnClassPath)
            {
                Stream result = OpenClasspathResource(sampleFileName);
                if (result == null)
                {
                    throw new Exception("specified test sample file '" + sampleFileName
                            + "' not found on the classpath");
                }
                //			System.out.println("Opening cp: " + sampleFileName);
                // wrap to avoid temp warning method about auto-closing input stream
                return new NonSeekableStream(result);
            }
            if (_resolvedDataDir == "")
            {
                throw new Exception("Must set system property '"
                        + TEST_DATA_DIR_SYS_PROPERTY_NAME
                        + "' properly before running tests");
            }


            if (!File.Exists(_resolvedDataDir+sampleFileName))
            {
                throw new Exception("Sample file '" + sampleFileName
                        + "' not found in data dir '" + _resolvedDataDir + "'");
            }

            
            //		System.out.println("Opening " + f.GetAbsolutePath());
            try
            {
                return new FileStream(_resolvedDataDir+sampleFileName,FileMode.Open);
            }
            catch (FileNotFoundException)
            {
                throw;
            }
        }

        private static void Initialise()
        {
            String dataDirName = System.Configuration.ConfigurationSettings.AppSettings[TEST_DATA_DIR_SYS_PROPERTY_NAME];

            if(dataDirName=="")
                    throw new Exception("Must set system property '"
                            + TEST_DATA_DIR_SYS_PROPERTY_NAME
                            + "' before running tests");

            if (!Directory.Exists(dataDirName))
            {
                throw new IOException("Data dir '" + dataDirName
                        + "' specified by system property '"
                        + TEST_DATA_DIR_SYS_PROPERTY_NAME + "' does not exist");
            }
            _sampleDataIsAvaliableOnClassPath = true;
            _resolvedDataDir = dataDirName;
        }

        /**
         * Opens a test sample file from the 'data' sub-package of this class's package. 
         * @return <code>null</code> if the sample file is1 not deployed on the classpath.
         */
        private static Stream OpenClasspathResource(String sampleFileName)
        {
            FileStream file = new FileStream(System.Configuration.ConfigurationSettings.AppSettings["HSSF.testdata.path"] + sampleFileName, FileMode.Open);
            return file;
        }

        private class NonSeekableStream : Stream
        {

            private Stream _is;

            public NonSeekableStream(Stream is1)
            {
                _is = is1;
            }

            public int Read()
            {
                return _is.ReadByte();
            }
            public override int Read(byte[] b, int off, int len)
            {
                return _is.Read(b, off, len);
            }
            public bool markSupported()
            {
                return false;
            }
            public override void Close()
            {
                _is.Close();
            }
            public override bool CanRead
            {
                get { return _is.CanRead; }
            }
            public override bool CanSeek
            {
                get { return false; }
            }
            public override bool CanWrite
            {
                get { return _is.CanWrite; }
            }
            public override long Length
            {
                get { return _is.Length; }
            }
            public override long Position
            {
                get { return _is.Position; }
                set { _is.Position = value; }
            }
            public override void Write(byte[] buffer, int offset, int count)
            {
                _is.Write(buffer, offset, count);
            }
            public override void Flush()
            {
                _is.Flush();
            }
            public override long Seek(long offset, SeekOrigin origin)
            {
                return _is.Seek(offset, origin);
            }
            public override void SetLength(long value)
            {
                _is.SetLength(value);
            }
        }

        public static HSSFWorkbook OpenSampleWorkbook(String sampleFileName)
        {
            try
            {
                return new HSSFWorkbook(OpenSampleFileStream(sampleFileName));
            }
            catch (IOException)
            {
                throw;
            }
        }
        /**
         * Writes a spReadsheet to a <tt>MemoryStream</tt> and Reads it back
         * from a <tt>ByteArrayStream</tt>.<p/>
         * Useful for verifying that the serialisation round trip
         */
        public static HSSFWorkbook WriteOutAndReadBack(HSSFWorkbook original)
        {

            try
            {
                MemoryStream baos = new MemoryStream(4096);
                original.Write(baos);
                return new HSSFWorkbook(baos);
            }
            catch (IOException)
            {
                throw;
            }
        }

        /**
         * @return byte array of sample file content from file found in standard hssf test data dir 
         */
        public static byte[] GetTestDataFileContent(String fileName)
        {
            MemoryStream bos = new MemoryStream();

            try
            {
                Stream fis = HSSFTestDataSamples.OpenSampleFileStream(fileName);

                byte[] buf = new byte[512];
                while (true)
                {
                    int bytesRead = fis.Read(buf,0,buf.Length);
                    if (bytesRead < 1)
                    {
                        break;
                    }
                    bos.Write(buf, 0, bytesRead);
                }
                fis.Close();
            }
            catch (IOException)
            {
                throw;
            }
            return bos.ToArray();
        }
    }
}