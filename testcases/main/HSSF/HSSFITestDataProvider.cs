using System;
using System.Collections.Generic;
using System.Text;
using TestCases.SS;
using NPOI.HSSF.UserModel;
using NPOI.SS;
using NPOI.SS.UserModel;
using System.IO;

namespace TestCases.HSSF
{
    public class HSSFITestDataProvider : ITestDataProvider
    {
        public IWorkbook OpenSampleWorkbook(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }
        public Stream OpenWorkbookStream(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleFileStream(sampleFileName);
        }

        public IWorkbook WriteOutAndReadBack(IWorkbook original)
        {
            if (!(original is HSSFWorkbook))
            {
                throw new ArgumentException("Expected an instance of HSSFWorkbook");
            }

            return HSSFTestDataSamples.WriteOutAndReadBack((HSSFWorkbook)original);
        }

        public IWorkbook CreateWorkbook()
        {
            return new HSSFWorkbook();
        }
        //************ SXSSF-specific methods ***************//
        public IWorkbook CreateWorkbook(int rowAccessWindowSize)
        {
            return CreateWorkbook();
        }
        public void TrackAllColumnsForAutosizing(ISheet sheet) { }
        //************ End SXSSF-specific methods ***************//
        public IFormulaEvaluator CreateFormulaEvaluator(IWorkbook wb)
        {
            return new HSSFFormulaEvaluator((HSSFWorkbook)wb);
        }
        public byte[] GetTestDataFileContent(String fileName)
        {
            return POIDataSamples.GetSpreadSheetInstance().ReadFile(fileName);
        }

        public SpreadsheetVersion GetSpreadsheetVersion()
        {
            return SpreadsheetVersion.EXCEL97;
        }

        public string StandardFileNameExtension
        {
            get { return "xls"; }
        }

        private HSSFITestDataProvider() { }
        private static HSSFITestDataProvider inst = new HSSFITestDataProvider();
        public static HSSFITestDataProvider Instance
        {
            get
            {
                return inst;
            }
        }
    }
}
