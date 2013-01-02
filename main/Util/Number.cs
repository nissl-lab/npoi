using System;

namespace NPOI.Util
{
    public class Number
    {
        private static Type BoolType = typeof(bool);
        private static Type CharType = typeof(char);
        private static Type IntPtrType = typeof(IntPtr);
        private static Type UIntPtrType = typeof(UIntPtr);
        private static Type DecimalType = typeof(decimal);
        public static bool IsNumber(object value)
        {
            if (value == null)
            {
                return false;
            }

            Type objType = value.GetType();
            if (objType.IsPrimitive || objType == DecimalType)
            {
                return (objType != BoolType &&
                        objType != CharType &&
                        objType != IntPtrType &&
                        objType != UIntPtrType);
            }
            //if (value is int) return true;
            //if (value is uint) return true;
            //if (value is long) return true;
            //if (value is ulong) return true;
            //if (value is sbyte) return true;
            //if (value is byte) return true;
            //if (value is short) return true;
            //if (value is ushort) return true;
            //if (value is float) return true;
            //if (value is double) return true;
            //if (value is decimal) return true;
            return false;
        }
        public static bool IsInteger(object value)
        {
            if (value == null)
            {
                return false;
            }
            if (value is int) return true;
            if (value is uint) return true;
            if (value is long) return true;
            if (value is ulong) return true;
            if (value is sbyte) return true;
            if (value is byte) return true;
            if (value is short) return true;
            if (value is ushort) return true;
            return false;
        }
    }
}