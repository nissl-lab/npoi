using System;
using System.Collections.Generic;
using System.Text;
using NPOI.Util;
using System.IO;

namespace TestCases.OpenXml4Net
{
    public class OpenXml4NetTestDataSamples
    {
        private static POIDataSamples _samples = POIDataSamples.GetOpenXML4JInstance();

        private OpenXml4NetTestDataSamples()
        {
            // no instances of this class
        }

        public static Stream OpenSampleStream(String sampleFileName)
        {
            return _samples.OpenResourceAsStream(sampleFileName);
        }
        public static String GetSampleFileName(String sampleFileName)
        {
            return new FileInfo(Path.Combine(_samples.ResolvedDataDir,sampleFileName)).FullName;
        }

        public static FileInfo GetSampleFile(String sampleFileName)
        {
            return new FileInfo(Path.Combine(_samples.ResolvedDataDir,sampleFileName));
        }

        public static FileInfo GetOutputFile(String outputFileName)
        {
            int dotPos = outputFileName.LastIndexOf('.');
            String suffix = outputFileName.Substring(dotPos);
            string path= TempFile.GetTempFilePath(outputFileName.Substring(0,dotPos), suffix);
            FileStream fs = File.Create(path);
            fs.Close();
            return new FileInfo(path);
        }


        public static Stream OpenComplianceSampleStream(String sampleFileName)
        {
            return _samples.OpenResourceAsStream(sampleFileName);
        }
    }
}
