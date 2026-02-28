namespace NPOI.SS.UserModel
{
    public interface IColor
    {
        short Indexed { get; }
        byte[] RGB { get; }
    }
}
