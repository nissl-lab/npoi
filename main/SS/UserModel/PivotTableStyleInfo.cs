namespace NPOI.SS.UserModel
{
    public interface IPivotTableStyleInfo : ITableStyleInfo
    {
        /// <summary>
        /// return true if column headers should be visible
        /// </summary>
        bool IsShowColumnHeaders { get; set; }

        /// <summary>
        /// return true if row headers should be visible
        /// </summary>
        bool IsShowRowHeaders { get; set; }
    }
}
