namespace NPOI.Util
{
    public static class Operator
    {
        public static int UnsignedRightShift(int operand,int val)
        {
            if (operand > 0)
                return operand >> val;
            else
                return (int)(((uint)operand) >> val);
        }
        public static long UnsignedRightShift(long operand, int val)
        {
            if (operand > 0)
                return operand >> val;
            else
                return (long)(((ulong)operand) >> val);
        }
        public static short UnsignedRightShift(short operand, int val)
        {
            
            if (operand > 0)
                return (short)(operand >> val);
            else
                return (short)(((ushort)operand) >> val);
        }
        public static sbyte UnsignedRightShift(sbyte operand, int val)
        {

            if (operand > 0)
                return (sbyte)(operand >> val);
            else
                return (sbyte)(((byte)operand) >> val);
        }
    }
}
