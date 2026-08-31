namespace NPOI.XWPF.UserModel
{
    /// <summary>
    /// The width types for tables and table cells. Table width can be specified as "auto" (AUTO),
    /// an absolute value in 20ths of a point (DXA), or as a percentage (PCT).
    /// </summary>
    public enum TableWidthType
    {
        AUTO = 3,
        DXA = 2,
        NIL = 0,
        PCT = 1
    }
}