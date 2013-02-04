
namespace NPOI.SS.UserModel
{
    public enum BorderDiagonal
    {
        /// <summary>
        /// No diagional border
        /// </summary>
        NONE=0,
        /// <summary>
        /// Backward diagional border, from left-top to right-bottom
        /// </summary>
        BACKWARD = 1,
        /// <summary>
        /// Forward diagional border, from right-top to left-bottom
        /// </summary>
        FORWARD=2,
        /// <summary>
        /// Both forward and backward diagional border
        /// </summary>
        BOTH=3
    }
}
