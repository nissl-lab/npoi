using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;

namespace TestCases
{
    public class POIDataSamples
    {
        /**
 * Name of the system property that defined path to the test data.
 */
        public static String TEST_PROPERTY = "POI.testdata.path";

        private static POIDataSamples _instSlideshow;
        private static POIDataSamples _instSpreadsheet;
        private static POIDataSamples _instDocument;
        private static POIDataSamples _instDiagram;
        private static POIDataSamples _instOpenXml4Net;
        private static POIDataSamples _instPOIFS;
        private static POIDataSamples _instDDF;
        private static POIDataSamples _instHPSF;
        private static POIDataSamples _instHPBF;
        private static POIDataSamples _instHSMF;

        private string _resolvedDataDir;
        /** <code>true</code> if standard system propery is not set,
         * but the data is available on the test runtime classpath */
        private bool _sampleDataIsAvaliableOnClassPath;
        private String _moduleDir;

        /**
 *
 * @param moduleDir   the name of the directory containing the test files
 */
        private POIDataSamples(String moduleDir)
        {
            _moduleDir = moduleDir;
            Initialise();
        }

        public static POIDataSamples GetSpreadSheetInstance()
        {
            if (_instSpreadsheet == null) _instSpreadsheet = new POIDataSamples("spreadsheet");
            return _instSpreadsheet;
        }

        public static POIDataSamples GetDocumentInstance()
        {
            if (_instDocument == null) _instDocument = new POIDataSamples("document");
            return _instDocument;
        }

        public static POIDataSamples GetSlideShowInstance()
        {
            if (_instSlideshow == null) _instSlideshow = new POIDataSamples("slideshow");
            return _instSlideshow;
        }

        public static POIDataSamples GetDiagramInstance()
        {
            if (_instOpenXml4Net == null) _instOpenXml4Net = new POIDataSamples("diagram");
            return _instOpenXml4Net;
        }

        public static POIDataSamples GetOpenXml4NetInstance()
        {
            if (_instDiagram == null) _instDiagram = new POIDataSamples("openxml4j");
            return _instDiagram;
        }

        public static POIDataSamples GetPOIFSInstance()
        {
            if (_instPOIFS == null) _instPOIFS = new POIDataSamples("poifs");
            return _instPOIFS;
        }

        public static POIDataSamples GetDDFInstance()
        {
            if (_instDDF == null) _instDDF = new POIDataSamples("ddf");
            return _instDDF;
        }

        public static POIDataSamples GetHPSFInstance()
        {
            if (_instHPSF == null) _instHPSF = new POIDataSamples("hpsf");
            return _instHPSF;
        }

        public static POIDataSamples GetPublisherInstance()
        {
            if (_instHPBF == null) _instHPBF = new POIDataSamples("publisher");
            return _instHPBF;
        }

        public static POIDataSamples GetHSMFInstance()
        {
            if (_instHSMF == null) _instHSMF = new POIDataSamples("hsmf");
            return _instHSMF;
        }

        /**
 * Opens a test sample file from the 'data' sub-package of this class's package. 
 * @return <c>null</c> if the sample file is1 not deployed on the classpath.
 */
        private Stream OpenClasspathResource(String sampleFileName)
        {
            FileStream file = new FileStream(_resolvedDataDir + sampleFileName, FileMode.Open, FileAccess.Read);
            return file;
        }

        private void Initialise()
        {
            String dataDirName = System.Configuration.ConfigurationSettings.AppSettings[TEST_PROPERTY];

            if (dataDirName == null)
                throw new Exception("Must set system property '"
                        + TEST_PROPERTY
                        + "' before running tests");

            string dataDir = string.Format(@"{0}\{1}\", dataDirName, _moduleDir);
            if (!Directory.Exists(dataDir))
            {
                throw new IOException("Data dir '" + dataDirName + "\\" + _moduleDir
                        + "' specified by system property '"
                        + TEST_PROPERTY + "' does not exist");
            }
            //if (!File.Exists(dataDirName + "SampleSS.xls"))
            //{
            //    throw new IOException(dataDirName + "SampleSS.xls does not exist");
            //}

            _sampleDataIsAvaliableOnClassPath = true;
            _resolvedDataDir = dataDir;
        }

        public string ResolvedDataDir
        {
            get { return _resolvedDataDir; }
        }

        /**
 * Opens a sample file from the standard HSSF test data directory
 * 
 * @return an Open <tt>Stream</tt> for the specified sample file
 */
        public Stream OpenResourceAsStream(String sampleFileName)
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
                        + TEST_PROPERTY
                        + "' properly before running tests");
            }


            if (!File.Exists(_resolvedDataDir + sampleFileName))
            {
                throw new Exception("Sample file '" + sampleFileName
                        + "' not found in data dir '" + _resolvedDataDir + "'");
            }


            //		System.out.println("Opening " + f.GetAbsolutePath());
            try
            {
                return new FileStream(_resolvedDataDir + sampleFileName, FileMode.Open, FileAccess.Read);
            }
            catch (FileNotFoundException)
            {
                throw;
            }
        }

        public FileInfo GetFileInfo(string sampleFileName)
        {
            string path = _resolvedDataDir + sampleFileName;
            if (!File.Exists(path))
            {
                throw new Exception("Sample file '" + sampleFileName
                        + "' not found in data dir '" + _resolvedDataDir + "'");
            }
            return new FileInfo(path);
        }

        /**
         *
         * @param sampleFileName    the name of the test file
         * @return
         * @throws RuntimeException if the file was not found
         */
        public FileStream GetFile(String sampleFileName)
        {
            string path=_resolvedDataDir+sampleFileName;
            if (!File.Exists(path))
            {
                throw new Exception("Sample file '" + sampleFileName
                        + "' not found in data dir '" + _resolvedDataDir + "'");
            }
            //try
            //{
            //    if (sampleFileName.Length > 0 && !sampleFileName.Equals(f.getCanonicalFile().getName()))
            //    {
            //        throw new RuntimeException("File name is case-sensitive: requested '" + sampleFileName
            //                + "' but actual file is '" + f.getCanonicalFile().getName() + "'");
            //    }
            //}
            //catch (IOException e)
            //{
            //    throw new RuntimeException(e);
            //}
            return new FileStream(path,FileMode.OpenOrCreate,FileAccess.ReadWrite);
        }
        public string[] GetFiles()
        {
            return Directory.GetFiles(_resolvedDataDir);
        }
        public string[] GetFiles(string searchPattern)
        {
            return Directory.GetFiles(_resolvedDataDir, searchPattern);
        }
        /**
 * @return byte array of sample file content from file found in standard hssf test data dir 
 */
        public byte[] ReadFile(String fileName)
        {
            MemoryStream bos = new MemoryStream();

            try
            {
                Stream fis = OpenResourceAsStream(fileName);

                byte[] buf = new byte[512];
                while (true)
                {
                    int bytesRead = fis.Read(buf, 0, buf.Length);
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
    }
}
