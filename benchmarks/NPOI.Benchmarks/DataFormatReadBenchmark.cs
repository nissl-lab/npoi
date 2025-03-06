using BenchmarkDotNet.Attributes;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace NPOI.Benchmarks;

[SimpleJob]
public class DataFormatReadBenchmark
{
    private const string dateFormat = "yyyy/MM/dd";
    private const string timeFormat = "HH:mm:ss";

    private HSSFWorkbook workbook;
    private ISheet sheet1;

    [Params(1_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        workbook = new HSSFWorkbook();
        Write(workbook);
    }

    private void Write(HSSFWorkbook wb)
    {
        var time = DateTime.UtcNow.AddYears(-1);

        IDataFormat df = wb.CreateDataFormat();
        var dateStyle = wb.CreateCellStyle();
        dateStyle.DataFormat = df.GetFormat(dateFormat);
        var timeStyle = wb.CreateCellStyle();
        timeStyle.DataFormat = df.GetFormat(timeFormat);

        ISheet sheet = wb.CreateSheet();

        for(var i = 0; i<RowCount; i++)
        {
            IRow row = sheet.CreateRow(i);

            ICell cellA = row.CreateCell(0);
            cellA.SetCellValue(time);
            cellA.CellStyle = dateStyle;

            ICell cellB = row.CreateCell(1);
            cellB.SetCellValue(time);
            cellB.CellStyle = timeStyle;

            time = time.AddHours(1);
        }
    }

    [Benchmark]
    public void Read()
    {
        DataFormatter df = new DataFormatter();
        ISheet sheet = workbook.GetSheetAt(0);
        var numRows = sheet.LastRowNum;
        for(var r = 0; r<=numRows; r++)
        {
            IRow row = sheet.GetRow(r);
            var numCells = row.LastCellNum;
            var readRow = new string[numCells];
            for(var c = 0; c<numCells; c++)
            {
                ICell cell = row.GetCell(c);
                if(cell is null) continue;
                readRow[c] = df.FormatCellValue(cell);
            }
        }
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        workbook.Dispose();
    }
}