namespace TestCases.POIFS.Crypt
{
    using NPOI.HSSF.Record.Crypto;
    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using System;
    using System.Collections.Generic;
    using System.IO;

    [TestFixture]
    public class TestPasswordEncryption
    {
        // 97 Binary file (xls)
        [TestCase("97_plain.xls", "97_plain_protected.xls", true)] // plain -> protect (default)
        [TestCase("97_password.xls", "97_password_unprotected.xls", false, false)] // protect -> plain
        [TestCase("97_rc4cryptoapi_password.xls", "97_rc4cryptoapi_password_unprotected.xls", false, false)] // protect -> plain
        [TestCase("97_password.xls", "97_password_protected.xls", false)] // protect -> protect
        [TestCase("97_rc4cryptoapi_password.xls", "97_rc4cryptoapi_password_protected.xls", false)] // protect -> protect
        //2007 Office Open XML file (xlsx)
        [TestCase("2007_plain.xlsx", "2007_plain_protected.xlsx", true)] // plain -> protect
        [TestCase("2007_password.xlsx", "2007_password_unprotected.xlsx", false, false)] // protect -> plain
        [TestCase("2007_password.xlsx", "2007_password_protected.xlsx", false)] // protect -> protect
        //2007 Macro enabled file (xlsm)
        [TestCase("2007macro_plain.xlsm", "2007macro_plain_protected.xlsm", true)] // plain -> protect
        [TestCase("2007macro_password.xlsm", "2007macro_password_unprotected.xlsm", false, false)] // protect -> plain
        [TestCase("2007macro_password.xlsm", "2007macro_password_protected.xlsm", false)] // protect -> protect
        public void TestWriteAndReadPasswordEncryption(
                  string inputFileName,
                  string outputFileName,
                  bool isPlain = false,
                  bool protectOutput = true,
                  string openPassword = "Password1234_")
        {
            FileInfo outputFile = TempFile.CreateTempFile(Path.GetFileNameWithoutExtension(outputFileName), Path.GetExtension(outputFileName));

            // Determine passwords
            string sourcePassword = isPlain ? null : openPassword;
            string outputPassword = protectOutput ? "new" + openPassword : null;

            // Load (decrypt if needed)
            Stream is1 = POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream(inputFileName);
            using var workbook = LoadWorkbook(is1, sourcePassword);

            // Save (encrypt if needed)
            using(var fs = new FileStream(outputFile.FullName, FileMode.Create, FileAccess.Write))
            {
                SaveWorkbook(workbook, fs, outputPassword);
            }

            // validate saved file content
            AssertSavedFileByReadWithNewPassword(outputFile.FullName, outputPassword);
        }

        private void SaveWorkbook(IWorkbook workbook, Stream outStream, string password = null)
        {
            if(string.IsNullOrEmpty(password))
            {
                var leaveStreamOpen = outStream is not FileStream;
                workbook.Write(outStream, leaveStreamOpen);
            }
            else
            {
                if(workbook is HSSFWorkbook hssfWorkbook)
                {
                    Biff8EncryptionKey.CurrentUserPassword = password;
                    try
                    {
                        workbook.Write(outStream);
                    }
                    finally
                    {
                        //Make sure to clear the password after saving to avoid affecting other operations
                        Biff8EncryptionKey.CurrentUserPassword = null;
                    }
                }
                else if(workbook is XSSFWorkbook xssfWorkbook)
                {
                    // Set up encryption
                    var fs = new POIFSFileSystem();
                    var info = new EncryptionInfo(EncryptionMode.Agile);
                    var encryptor = info.Encryptor;
                    encryptor.ConfirmPassword(password);

                    // Save workbook to encrypted output stream
                    using(var outputStream = encryptor.GetDataStream(fs))
                    {
                        xssfWorkbook.Write(outputStream);
                    }

                    // Write out the encrypted version
                    fs.WriteFileSystem(outStream);
                }
                else
                {
                    throw new NotSupportedException("Workbook type not supported for password protection.");
                }
            }
        }

        private IWorkbook LoadWorkbook(Stream stream, string openPassword)
        {
            if(stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if(stream.CanSeek)
            {
                stream.Seek(0, SeekOrigin.Begin);
            }

            return WorkbookFactory.Create(stream, openPassword);
        }

        private void AssertSavedFileByReadWithNewPassword(string filePath, string newPassword)
        {
            // Re-open with NPOI
            using var protectedStream = File.OpenRead(filePath);
            using var protectedWb = LoadWorkbook(protectedStream, newPassword);
            for(int sheetNo = 0; sheetNo < protectedWb.NumberOfSheets; sheetNo++)
            {
                var protectedSheet = protectedWb.GetSheetAt(sheetNo);
                var data = ReadSheetData(protectedSheet, trimValues: true, includeEmptyTrailingCells: false);

                ClassicAssert.AreEqual(protectedSheet.LastRowNum, data.Count - 1);
            }
        }

        private List<List<string>> ReadSheetData(ISheet sheet, bool trimValues = true, bool includeEmptyTrailingCells = false)
        {
            var result = new List<List<string>>();
            if(sheet == null)
                return result;

            // Determine max column count based on first non-null row (could adapt to scan all if needed)
            int maxColumns = 0;
            for(int r = sheet.FirstRowNum; r <= sheet.LastRowNum && r < sheet.FirstRowNum + 50; r++)
            {
                var probe = sheet.GetRow(r);
                if(probe == null)
                    continue;
                if(probe.LastCellNum > maxColumns)
                    maxColumns = probe.LastCellNum;
                if(maxColumns > 0)
                    break;
            }

            // Fallback: scan all rows if still zero
            if(maxColumns == 0)
            {
                for(int r = sheet.FirstRowNum; r <= sheet.LastRowNum; r++)
                {
                    var probe = sheet.GetRow(r);
                    if(probe == null)
                        continue;
                    if(probe.LastCellNum > maxColumns)
                        maxColumns = probe.LastCellNum;
                }
            }

            for(int r = sheet.FirstRowNum; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if(row == null)
                {
                    if(includeEmptyTrailingCells)
                        result.Add(new List<string>(new string[maxColumns]));
                    else
                        result.Add(new List<string>());
                    continue;
                }

                // If we don't want fixed width, just iterate existing cells
                if(!includeEmptyTrailingCells && maxColumns == 0)
                    maxColumns = row.LastCellNum;

                var rowData = new List<string>(maxColumns);

                for(int c = 0; c < maxColumns; c++)
                {
                    var cell = row.GetCell(c);
                    if (cell == null)
                    {
                        rowData.Add("");
                        continue;
                    }
                    var value = GetValue(cell);
                    if(trimValues && value is string strValue)
                        strValue = strValue.Trim();
                    rowData.Add(value?.ToString() ?? "");
                }

                if(!includeEmptyTrailingCells)
                {
                    // Optionally trim trailing empties for a cleaner structure
                    for(int i = rowData.Count - 1; i >= 0; i--)
                    {
                        if(rowData[i].Length == 0)
                            rowData.RemoveAt(i);
                        else
                            break;
                    }
                }

                result.Add(rowData);
            }

            return result;
        }

        private object GetValue(NPOI.SS.UserModel.ICell cell)
        {
            return cell.CellType switch {
                CellType.Blank => null,
                CellType.Boolean => cell.BooleanCellValue,
                CellType.Numeric => DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue : cell.NumericCellValue,
                CellType.String => cell.StringCellValue,
                CellType.Formula => cell.CachedFormulaResultType switch {
                    CellType.Boolean => cell.BooleanCellValue,
                    CellType.Numeric => DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue : cell.NumericCellValue,
                    CellType.String => cell.StringCellValue,
                    _ => null
                },
                _ => null
            };
        }
    }
}