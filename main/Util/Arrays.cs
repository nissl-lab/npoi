
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System;
using System.Collections;
using System.Text;


namespace NPOI.Util
{
    public class Arrays
    {
        /// <summary>
        /// Fills the specified array.
        /// </summary>
        /// <param name="array">The array.</param>
        /// <param name="defaultValue">The default value.</param>
        public static void Fill(byte[] array,byte defaultValue)
        {
            for (int i = 0; i < array.Length; i++)
            {
                array[i] = defaultValue;
            }
        }
        public static void Fill(char[] array, char defaultValue)
        {
            for (int i = 0; i < array.Length; i++)
            {
                array[i] = defaultValue;
            }
        }
        public static void Fill<T>(T[] array, T defaultValue)
        {
            for (int i = 0; i < array.Length; i++)
            {
                array[i] = defaultValue;
            }
        }

        /// <summary>
        /// Assigns the specified byte value to each element of the specified
        /// range of the specified array of bytes.  The range to be filled
        /// extends from index <tt>fromIndex</tt>, inclusive, to index
        /// <tt>toIndex</tt>, exclusive.  (If <tt>fromIndex==toIndex</tt>, the
        /// range to be filled is empty.)
        /// </summary>
        /// <param name="a">the array to be filled</param>
        /// <param name="fromIndex">the index of the first element (inclusive) to be filled with the specified value</param>
        /// <param name="toIndex">the index of the last element (exclusive) to be filled with the specified value</param>
        /// <param name="val">the value to be stored in all elements of the array</param>
        /// <exception cref="System.ArgumentException">if <c>fromIndex &gt; toIndex</c></exception>
        /// <exception cref="System.IndexOutOfRangeException"> if <c>fromIndex &lt; 0</c> or <c>toIndex &gt; a.length</c></exception>
        public static void Fill(byte[] a, int fromIndex, int toIndex, byte val)
        {
            RangeCheck(a.Length, fromIndex, toIndex);
            for (int i = fromIndex; i < toIndex; i++)
                a[i] = val;
        }

        /// <summary>
        /// Checks that {@code fromIndex} and {@code toIndex} are in
        /// the range and throws an appropriate exception, if they aren't.
        /// </summary>
        /// <param name="length"></param>
        /// <param name="fromIndex"></param>
        /// <param name="toIndex"></param>
        private static void RangeCheck(int length, int fromIndex, int toIndex)
        {
            if (fromIndex > toIndex)
            {
                throw new ArgumentException(
                    "fromIndex(" + fromIndex + ") > toIndex(" + toIndex + ")");
            }
            if (fromIndex < 0)
            {
                throw new IndexOutOfRangeException("fromIndex(" + fromIndex + ")");
            }
            if (toIndex > length)
            {
                throw new IndexOutOfRangeException( "toIndex(" + toIndex + ")");
            }
        }
        /// <summary>
        /// Convert Array to ArrayList
        /// </summary>
        /// <param name="arr">source array</param>
        /// <returns></returns>
        public static ArrayList AsList(Array arr)
        {
            if (arr.Length <= 0)
                return new ArrayList();
            ArrayList al = new ArrayList(arr.Length);
            for (int i = 0; i < arr.Length; i++)
            {
                al.Add(arr.GetValue(i));
            }
            return al;
        }
        /// <summary>
        /// Fills the specified array.
        /// </summary>
        /// <param name="array">The array.</param>
        /// <param name="defaultValue">The default value.</param>
        public static void Fill(int[] array, byte defaultValue)
        {
            for (int i = 0; i < array.Length; i++)
            {
                array[i] = defaultValue;
            }
        }

        /// <summary>
        /// Equals the specified a1.
        /// </summary>
        /// <param name="a1">The a1.</param>
        /// <param name="b1">The b1.</param>
        /// <returns></returns>
        public new static bool Equals(object a1, object b1)
        {
            if (a1 == null || b1 == null)
                return false;
            Array a = a1 as Array;
            Array b = b1 as Array;
            if (a.Length != b.Length)
                return false;

            for (int i = 0; i < a.Length; i++)
            {
                if (!a.GetValue(i).Equals(b.GetValue(i)))
                    return false;
            }
            return true;
        }
        /**
         * Returns <c>true</c> if the two specified arrays of Objects are
         * <i>equal</i> to one another.  The two arrays are considered equal if
         * both arrays contain the same number of elements, and all corresponding
         * pairs of elements in the two arrays are equal.  Two objects <c>e1</c>
         * and <c>e2</c> are considered <i>equal</i> if <c>(e1==null ? e2==null
         * : e1.equals(e2))</c>.  In other words, the two arrays are equal if
         * they contain the same elements in the same order.  Also, two array
         * references are considered equal if both are <c>null</c>.
         *
         * @param a one array to be tested for equality
         * @param a2 the other array to be tested for equality
         * @return <c>true</c> if the two arrays are equal
         */
        public static bool Equals(Object[] a, Object[] a2)
        {
            if (a == a2)
                return true;
            if (a == null || a2 == null)
                return false;

            int length = a.Length;
            if (a2.Length != length)
                return false;

            for (int i = 0; i < length; i++)
            {
                Object o1 = a[i];
                Object o2 = a2[i];
                if (!(o1 == null ? o2 == null : o1.Equals(o2)))
                    return false;
            }

            return true;
        }
        /// <summary>
        /// Moves a number of entries in an array to another point in the array, shifting those inbetween as required.
        /// </summary>
        /// <param name="array">The array to alter</param>
        /// <param name="moveFrom">The (0 based) index of the first entry to move</param>
        /// <param name="moveTo">The (0 based) index of the positition to move to</param>
        /// <param name="numToMove">The number of entries to move</param>
        public static void ArrayMoveWithin(Object[] array, int moveFrom, int moveTo, int numToMove)
        {
            // If we're not asked to do anything, return now
            if (numToMove <= 0) { return; }
            if (moveFrom == moveTo) { return; }

            // Check that the values supplied are valid
            if (moveFrom < 0 || moveFrom >= array.Length)
            {
                throw new ArgumentException("The moveFrom must be a valid array index");
            }
            if (moveTo < 0 || moveTo >= array.Length)
            {
                throw new ArgumentException("The moveTo must be a valid array index");
            }
            if (moveFrom + numToMove > array.Length)
            {
                throw new ArgumentException("Asked to move more entries than the array has");
            }
            if (moveTo + numToMove > array.Length)
            {
                throw new ArgumentException("Asked to move to a position that doesn't have enough space");
            }

            // Grab the bit to move 
            Object[] toMove = new Object[numToMove];
            Array.Copy(array, moveFrom, toMove, 0, numToMove);

            // Grab the bit to be shifted
            Object[] toShift;
            int shiftTo;
            if (moveFrom > moveTo)
            {
                // Moving to an earlier point in the array
                // Grab everything between the two points
                toShift = new Object[(moveFrom - moveTo)];
                Array.Copy(array, moveTo, toShift, 0, toShift.Length);
                shiftTo = moveTo + numToMove;
            }
            else
            {
                // Moving to a later point in the array
                // Grab everything from after the toMove block, to the new point
                toShift = new Object[(moveTo - moveFrom)];
                Array.Copy(array, moveFrom + numToMove, toShift, 0, toShift.Length);
                shiftTo = moveFrom;
            }

            // Copy the moved block to its new location
            Array.Copy(toMove, 0, array, moveTo, toMove.Length);

            // And copy the shifted block to the shifted location
            Array.Copy(toShift, 0, array, shiftTo, toShift.Length);


            // We're done - array will now have everything moved as required
        }

        /// <summary>
        ///  Copies the specified array, truncating or padding with zeros (if
        /// necessary) so the copy has the specified length. This method is temporary
        /// replace for Arrays.copyOf() until we start to require JDK 1.6.
        /// </summary>
        /// <param name="source">the array to be copied</param>
        /// <param name="newLength">the length of the copy to be returned</param>
        /// <returns>a copy of the original array, truncated or padded with zeros to obtain the specified length</returns>
        public static byte[] CopyOf(byte[] source, int newLength)
        {
            byte[] result = new byte[newLength];
            Array.Copy(source, 0, result, 0,
                    Math.Min(source.Length, newLength));
            return result;
        }

        internal static int[] CopyOfRange(int[] original, int from, int to)
        {
            int newLength = to - from;
            if (newLength < 0)
                throw new ArgumentException(from + " > " + to);
            int[] copy = new int[newLength];
            Array.Copy(original, from, copy, 0,
                             Math.Min(original.Length - from, newLength));
            return copy;
        }
        internal static byte[] CopyOfRange(byte[] original, int from, int to)
        {
            int newLength = to - from;
            if (newLength < 0)
                throw new ArgumentException(from + " > " + to);
            byte[] copy = new byte[newLength];
            Array.Copy(original, from, copy, 0,
                             Math.Min(original.Length - from, newLength));
            return copy;
        }
        /**
         * Returns a string representation of the contents of the specified array.
         * If the array contains other arrays as elements, they are converted to
         * strings by the {@link Object#toString} method inherited from
         * <tt>Object</tt>, which describes their <i>identities</i> rather than
         * their contents.
         *
         * <p>The value returned by this method is equal to the value that would
         * be returned by <tt>Arrays.asList(a).toString()</tt>, unless <tt>a</tt>
         * is <tt>null</tt>, in which case <tt>"null"</tt> is returned.</p>
         *
         * @param a the array whose string representation to return
         * @return a string representation of <tt>a</tt>
         * @see #deepToString(Object[])
         * @since 1.5
         */
        public static String ToString(Object[] a)
        {
            if (a == null)
                return "null";

            int iMax = a.Length - 1;
            if (iMax == -1)
                return "[]";

            StringBuilder b = new StringBuilder();
            b.Append('[');
            for (int i = 0; ; i++)
            {
                b.Append(a[i].ToString());
                if (i == iMax)
                    return b.Append(']').ToString();
                b.Append(", ");
            }
        }
    }
}
