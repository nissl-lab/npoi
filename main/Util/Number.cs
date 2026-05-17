using ExtendedNumerics;
using System;
using System.Collections;
using System.Numerics;

namespace NPOI.Util
{
    public class Number
    {
        public static int BitCount(int i)
        {
#if NET8_0_OR_GREATER
            return BitOperations.PopCount((uint)i);
#else
            BitArray bitArray = new BitArray(BitConverter.GetBytes(i));
            int count = 0;
            for (int idx = 0; idx < bitArray.Count; idx++)
            {
                if (bitArray.Get(idx))
                    count++;
            }
            return count;
            //i = i - ((i >>> 1) & 0x55555555);
            //i = (i & 0x33333333) + ((i >>> 2) & 0x33333333);
            //i = (i + (i >>> 4)) & 0x0f0f0f0f;
            //i = i + (i >>> 8);
            //i = i + (i >>> 16);
            //return i & 0x3f;
#endif

        }
        private static readonly Type BoolType = typeof(bool);
        private static readonly Type CharType = typeof(char);
        private static readonly Type IntPtrType = typeof(IntPtr);
        private static readonly Type UIntPtrType = typeof(UIntPtr);
        private static readonly Type DecimalType = typeof(decimal);
        private static readonly Type BigDecimalType = typeof(BigDecimal);
        public static bool IsNumber(object value)
        {
            if (value == null)
            {
                return false;
            }

            Type objType = value.GetType();
            if (objType.IsPrimitive || objType == DecimalType
                || objType == BigDecimalType)
            {
                return (objType != BoolType &&
                        objType != CharType &&
                        objType != IntPtrType &&
                        objType != UIntPtrType);
            }
            return false;
        }

        public static bool IsNumber(Type objType)
        {
            if (objType.IsPrimitive || objType == DecimalType)
            {
                return (objType != BoolType &&
                        objType != CharType &&
                        objType != IntPtrType &&
                        objType != UIntPtrType);
            }
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