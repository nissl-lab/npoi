
namespace NPOI.SS.UserModel
{
    public enum BorderDiagonal
    {
        /// <summary>
        /// No diagional border
        /// </summary>
        None = 0,
        /// <summary>
        /// Backward diagional border, from left-top to right-bottom
        /// </summary>
        Backward = 1,
        /// <summary>
        /// Forward diagional border, from right-top to left-bottom
        /// </summary>
        Forward = 2,
        /// <summary>
        /// Both forward and backward diagional border
        /// </summary>
        Both = 3
    }
}
