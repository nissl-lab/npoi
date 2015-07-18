using System;
using NPOI.SS.UserModel;
using NPOI.SS;

namespace TestCases.SS
{
    /// <summary>
    /// Encapsulates a provider of test data for common HSSF / XSSF tests.
    /// </summary>
    public interface ITestDataProvider
    {
        /// <summary>
        /// Override to provide HSSF / XSSF specific way for re-serialising a workbook
        /// </summary>
        /// <param name="wb">the workbook to re-serialize.</param>
        /// <returns>the re-serialized workbook</returns>
        IWorkbook WriteOutAndReadBack(IWorkbook wb);

        /// <summary>
        /// Override to provide way of loading HSSF / XSSF sample workbooks
        /// </summary>
        /// <param name="sampleFileName"> the file name to load.</param>
        /// <returns>an instance of Workbook loaded from the supplied file name</returns>
        IWorkbook OpenSampleWorkbook(String sampleFileName);

        /// <summary>
        /// Override to provide way of creating HSSF / XSSF workbooks
        /// </summary>
        /// <returns>an instance of Workbook</returns>
        IWorkbook CreateWorkbook();

        /// <summary>
        ///Opens a sample file from the standard HSSF test data directory
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>an open InputStream for the specified sample file</returns>
        byte[] GetTestDataFileContent(String fileName);

        SpreadsheetVersion GetSpreadsheetVersion();

        string StandardFileNameExtension { get; }

        /**
         * Creates the corresponding {@link FormulaEvaluator} for the
         * type of Workbook handled by this Provider. 
         *
         * @param wb The workbook to base the formula evaluator on.
         * @return A new instance of a matching type of formula evaluator. 
         */
        IFormulaEvaluator CreateFormulaEvaluator(IWorkbook wb);
    }
}
