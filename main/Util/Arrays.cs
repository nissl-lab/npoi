
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
using System.Collections.Generic;
using NPOI.Util.Collections;


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
         * Returns a hash code based on the contents of the specified array.
         * For any two <tt>long</tt> arrays <tt>a</tt> and <tt>b</tt>
         * such that <tt>Arrays.Equals(a, b)</tt>, it is also the case that
         * <tt>Arrays.HashCode(a) == Arrays.HashCode(b)</tt>.
         *
         * The value returned by this method is the same value that would be
         * obtained by invoking the {@link List#hashCode() <tt>hashCode</tt>}
         * method on a {@link List} Containing a sequence of {@link Long}
         * instances representing the elements of <tt>a</tt> in the same order.
         * If <tt>a</tt> is <tt>null</tt>, this method returns 0.
         *
         * @param a the array whose hash value to compute
         * @return a content-based hash code for <tt>a</tt>
         * @since 1.5
         */
        public static int HashCode(long[] a) {
            if (a == null)
                return 0;

            int result = 1;
            foreach (long element in a) {
                int elementHash = (int)(element ^ (Operator.UnsignedRightShift(element , 32)));
                result = 31 * result + elementHash;
            }

            return result;
        }

        /**
         * Returns a hash code based on the contents of the specified array.
         * For any two non-null <tt>int</tt> arrays <tt>a</tt> and <tt>b</tt>
         * such that <tt>Arrays.Equals(a, b)</tt>, it is also the case that
         * <tt>Arrays.HashCode(a) == Arrays.HashCode(b)</tt>.
         *
         * The value returned by this method is the same value that would be
         * obtained by invoking the {@link List#hashCode() <tt>hashCode</tt>}
         * method on a {@link List} Containing a sequence of {@link int}
         * instances representing the elements of <tt>a</tt> in the same order.
         * If <tt>a</tt> is <tt>null</tt>, this method returns 0.
         *
         * @param a the array whose hash value to compute
         * @return a content-based hash code for <tt>a</tt>
         * @since 1.5
         */
        public static int HashCode(int[] a) {
            if (a == null)
                return 0;

            int result = 1;
            foreach (int element in a)
                result = 31 * result + element;

            return result;
        }

        /**
         * Returns a hash code based on the contents of the specified array.
         * For any two <tt>short</tt> arrays <tt>a</tt> and <tt>b</tt>
         * such that <tt>Arrays.Equals(a, b)</tt>, it is also the case that
         * <tt>Arrays.HashCode(a) == Arrays.HashCode(b)</tt>.
         *
         * The value returned by this method is the same value that would be
         * obtained by invoking the {@link List#hashCode() <tt>hashCode</tt>}
         * method on a {@link List} Containing a sequence of {@link short}
         * instances representing the elements of <tt>a</tt> in the same order.
         * If <tt>a</tt> is <tt>null</tt>, this method returns 0.
         *
         * @param a the array whose hash value to compute
         * @return a content-based hash code for <tt>a</tt>
         * @since 1.5
         */
        public static int HashCode(short[] a) {
            if (a == null)
                return 0;

            int result = 1;
            foreach (short element in a)
                result = 31 * result + element;

            return result;
        }

        /**
         * Returns a hash code based on the contents of the specified array.
         * For any two <tt>char</tt> arrays <tt>a</tt> and <tt>b</tt>
         * such that <tt>Arrays.Equals(a, b)</tt>, it is also the case that
         * <tt>Arrays.HashCode(a) == Arrays.HashCode(b)</tt>.
         *
         * The value returned by this method is the same value that would be
         * obtained by invoking the {@link List#hashCode() <tt>hashCode</tt>}
         * method on a {@link List} Containing a sequence of {@link Character}
         * instances representing the elements of <tt>a</tt> in the same order.
         * If <tt>a</tt> is <tt>null</tt>, this method returns 0.
         *
         * @param a the array whose hash value to compute
         * @return a content-based hash code for <tt>a</tt>
         * @since 1.5
         */
        public static int HashCode(char[] a) {
            if (a == null)
                return 0;

            int result = 1;
            foreach (char element in a)
                result = 31 * result + element;

            return result;
        }

        /**
         * Returns a hash code based on the contents of the specified array.
         * For any two <tt>byte</tt> arrays <tt>a</tt> and <tt>b</tt>
         * such that <tt>Arrays.Equals(a, b)</tt>, it is also the case that
         * <tt>Arrays.HashCode(a) == Arrays.HashCode(b)</tt>.
         *
         * The value returned by this method is the same value that would be
         * obtained by invoking the {@link List#hashCode() <tt>hashCode</tt>}
         * method on a {@link List} Containing a sequence of {@link Byte}
         * instances representing the elements of <tt>a</tt> in the same order.
         * If <tt>a</tt> is <tt>null</tt>, this method returns 0.
         *
         * @param a the array whose hash value to compute
         * @return a content-based hash code for <tt>a</tt>
         * @since 1.5
         */
        public static int HashCode(byte[] a) {
            if (a == null)
                return 0;

            int result = 1;
            foreach (byte element in a)
                result = 31 * result + element;

            return result;
        }

        /**
         * Returns a hash code based on the contents of the specified array.
         * For any two <tt>bool</tt> arrays <tt>a</tt> and <tt>b</tt>
         * such that <tt>Arrays.Equals(a, b)</tt>, it is also the case that
         * <tt>Arrays.HashCode(a) == Arrays.HashCode(b)</tt>.
         *
         * The value returned by this method is the same value that would be
         * obtained by invoking the {@link List#hashCode() <tt>hashCode</tt>}
         * method on a {@link List} Containing a sequence of {@link Boolean}
         * instances representing the elements of <tt>a</tt> in the same order.
         * If <tt>a</tt> is <tt>null</tt>, this method returns 0.
         *
         * @param a the array whose hash value to compute
         * @return a content-based hash code for <tt>a</tt>
         * @since 1.5
         */
        public static int HashCode(bool[] a) {
            if (a == null)
                return 0;

            int result = 1;
            foreach (bool element in a)
                result = 31 * result + (element ? 1231 : 1237);

            return result;
        }

        /**
         * Returns a hash code based on the contents of the specified array.
         * For any two <tt>float</tt> arrays <tt>a</tt> and <tt>b</tt>
         * such that <tt>Arrays.Equals(a, b)</tt>, it is also the case that
         * <tt>Arrays.HashCode(a) == Arrays.HashCode(b)</tt>.
         *
         * The value returned by this method is the same value that would be
         * obtained by invoking the {@link List#hashCode() <tt>hashCode</tt>}
         * method on a {@link List} Containing a sequence of {@link Float}
         * instances representing the elements of <tt>a</tt> in the same order.
         * If <tt>a</tt> is <tt>null</tt>, this method returns 0.
         *
         * @param a the array whose hash value to compute
         * @return a content-based hash code for <tt>a</tt>
         * @since 1.5
         */
        public static int HashCode(float[] a) {
            if (a == null)
                return 0;

            int result = 1;
            foreach (float element in a)
            {
                result = 31 * result + BitConverter.ToInt32(BitConverter.GetBytes(element), 0);
            }

            return result;
        }

        /**
         * Returns a hash code based on the contents of the specified array.
         * For any two <tt>double</tt> arrays <tt>a</tt> and <tt>b</tt>
         * such that <tt>Arrays.Equals(a, b)</tt>, it is also the case that
         * <tt>Arrays.HashCode(a) == Arrays.HashCode(b)</tt>.
         *
         * The value returned by this method is the same value that would be
         * obtained by invoking the {@link List#hashCode() <tt>hashCode</tt>}
         * method on a {@link List} Containing a sequence of {@link Double}
         * instances representing the elements of <tt>a</tt> in the same order.
         * If <tt>a</tt> is <tt>null</tt>, this method returns 0.
         *
         * @param a the array whose hash value to compute
         * @return a content-based hash code for <tt>a</tt>
         * @since 1.5
         */
        public static int HashCode(double[] a) {
            if (a == null)
                return 0;

            int result = 1;
            foreach (double element in a) {
                long bits = BitConverter.DoubleToInt64Bits(element);
                result = 31 * result + (int)(bits ^ (Operator.UnsignedRightShift(bits, 32)));
            }
            return result;
        }

        /**
         * Returns a hash code based on the contents of the specified array.  If
         * the array Contains other arrays as elements, the hash code is based on
         * their identities rather than their contents.  It is therefore
         * acceptable to invoke this method on an array that Contains itself as an
         * element,  either directly or indirectly through one or more levels of
         * arrays.
         *
         * For any two arrays <tt>a</tt> and <tt>b</tt> such that
         * <tt>Arrays.Equals(a, b)</tt>, it is also the case that
         * <tt>Arrays.HashCode(a) == Arrays.HashCode(b)</tt>.
         *
         * The value returned by this method is equal to the value that would
         * be returned by <tt>Arrays.AsList(a).HashCode()</tt>, unless <tt>a</tt>
         * is <tt>null</tt>, in which case <tt>0</tt> is returned.
         *
         * @param a the array whose content-based hash code to compute
         * @return a content-based hash code for <tt>a</tt>
         * @see #deepHashCode(Object[])
         * @since 1.5
         */
        public static int HashCode(Object[] a) {
            if (a == null)
                return 0;

            int result = 1;

            foreach (Object element in a)
                result = 31 * result + (element == null ? 0 : element.GetHashCode());

            return result;
        }

        /**
         * Returns a hash code based on the "deep contents" of the specified
         * array.  If the array Contains other arrays as elements, the
         * hash code is based on their contents and so on, ad infInitum.
         * It is therefore unacceptable to invoke this method on an array that
         * Contains itself as an element, either directly or indirectly through
         * one or more levels of arrays.  The behavior of such an invocation is
         * undefined.
         *
         * For any two arrays <tt>a</tt> and <tt>b</tt> such that
         * <tt>Arrays.DeepEquals(a, b)</tt>, it is also the case that
         * <tt>Arrays.DeepHashCode(a) == Arrays.DeepHashCode(b)</tt>.
         *
         * The computation of the value returned by this method is similar to
         * that of the value returned by {@link List#hashCode()} on a list
         * Containing the same elements as <tt>a</tt> in the same order, with one
         * difference: If an element <tt>e</tt> of <tt>a</tt> is itself an array,
         * its hash code is computed not by calling <tt>e.HashCode()</tt>, but as
         * by calling the appropriate overloading of <tt>Arrays.HashCode(e)</tt>
         * if <tt>e</tt> is an array of a primitive type, or as by calling
         * <tt>Arrays.DeepHashCode(e)</tt> recursively if <tt>e</tt> is an array
         * of a reference type.  If <tt>a</tt> is <tt>null</tt>, this method
         * returns 0.
         *
         * @param a the array whose deep-content-based hash code to compute
         * @return a deep-content-based hash code for <tt>a</tt>
         * @see #hashCode(Object[])
         * @since 1.5
         */
        public static int DeepHashCode(Object[] a) {
            if (a == null)
                return 0;

            int result = 1;

            foreach (Object element in a) {
                int elementHash = 0;
                if (element is Object[])
                    elementHash = DeepHashCode((Object[]) element);
                else if (element is byte[])
                    elementHash = HashCode((byte[]) element);
                else if (element is short[])
                    elementHash = HashCode((short[]) element);
                else if (element is int[])
                    elementHash = HashCode((int[]) element);
                else if (element is long[])
                    elementHash = HashCode((long[]) element);
                else if (element is char[])
                    elementHash = HashCode((char[]) element);
                else if (element is float[])
                    elementHash = HashCode((float[]) element);
                else if (element is double[])
                    elementHash = HashCode((double[]) element);
                else if (element is bool[])
                    elementHash = HashCode((bool[]) element);
                else if (element != null)
                    elementHash = element.GetHashCode();

                result = 31 * result + elementHash;
            }

            return result;
        }

        /**
         * Returns <tt>true</tt> if the two specified arrays are <i>deeply
         * Equal</i> to one another.  Unlike the {@link #Equals(Object[],Object[])}
         * method, this method is appropriate for use with nested arrays of
         * arbitrary depth.
         *
         * Two array references are considered deeply equal if both
         * are <tt>null</tt>, or if they refer to arrays that contain the same
         * number of elements and all corresponding pairs of elements in the two
         * arrays are deeply Equal.
         *
         * Two possibly <tt>null</tt> elements <tt>e1</tt> and <tt>e2</tt> are
         * deeply equal if any of the following conditions hold:
         * <ul>
         *    <li> <tt>e1</tt> and <tt>e2</tt> are both arrays of object reference
         *         types, and <tt>Arrays.DeepEquals(e1, e2) would return true</tt></li>
         *    <li> <tt>e1</tt> and <tt>e2</tt> are arrays of the same primitive
         *         type, and the appropriate overloading of
         *         <tt>Arrays.Equals(e1, e2)</tt> would return true.</li>
         *    <li> <tt>e1 == e2</tt></li>
         *    <li> <tt>e1.Equals(e2)</tt> would return true.</li>
         * </ul>
         * Note that this defInition permits <tt>null</tt> elements at any depth.
         *
         * If either of the specified arrays contain themselves as elements
         * either directly or indirectly through one or more levels of arrays,
         * the behavior of this method is undefined.
         *
         * @param a1 one array to be tested for Equality
         * @param a2 the other array to be tested for Equality
         * @return <tt>true</tt> if the two arrays are equal
         * @see #Equals(Object[],Object[])
         * @see Objects#deepEquals(Object, Object)
         * @since 1.5
         */
        public static bool DeepEquals(Object[] a1, Object[] a2) {
            if (a1 == a2)
                return true;
            if (a1 == null || a2==null)
                return false;
            int length = a1.Length;
            if (a2.Length != length)
                return false;

            for (int i = 0; i < length; i++) {
                Object e1 = a1[i];
                Object e2 = a2[i];

                if (e1 == e2)
                    continue;
                if (e1 == null)
                    return false;

                // Figure out whether the two elements are equal
                bool eq = DeepEquals0(e1, e2);

                if (!eq)
                    return false;
            }
            return true;
        }

        static bool DeepEquals0(Object e1, Object e2) {
            bool eq;
            if (e1 is Object[] && e2 is Object[])
                eq = DeepEquals ((Object[]) e1, (Object[]) e2);
            else if (e1 is byte[] && e2 is byte[])
                eq = Equals((byte[]) e1, (byte[]) e2);
            else if (e1 is short[] && e2 is short[])
                eq = Equals((short[]) e1, (short[]) e2);
            else if (e1 is int[] && e2 is int[])
                eq = Equals((int[]) e1, (int[]) e2);
            else if (e1 is long[] && e2 is long[])
                eq = Equals((long[]) e1, (long[]) e2);
            else if (e1 is char[] && e2 is char[])
                eq = Equals((char[]) e1, (char[]) e2);
            else if (e1 is float[] && e2 is float[])
                eq = Equals((float[]) e1, (float[]) e2);
            else if (e1 is double[] && e2 is double[])
                eq = Equals((double[]) e1, (double[]) e2);
            else if (e1 is bool[] && e2 is bool[])
                eq = Equals((bool[]) e1, (bool[]) e2);
            else
                eq = e1.Equals(e2);
            return eq;
        }

        /**
         * Returns a string representation of the contents of the specified array.
         * The string representation consists of a list of the array's elements,
         * enclosed in square brackets (<tt>"[]"</tt>).  Adjacent elements are
         * Separated by the characters <tt>", "</tt> (a comma followed by a
         * space).  Elements are Converted to strings as by
         * <tt>String.ValueOf(long)</tt>.  Returns <tt>"null"</tt> if <tt>a</tt>
         * is <tt>null</tt>.
         *
         * @param a the array whose string representation to return
         * @return a string representation of <tt>a</tt>
         * @since 1.5
         */
        public static String ToString(long[] a) {
            if (a == null)
                return "null";
            int iMax = a.Length - 1;
            if (iMax == -1)
                return "[]";

            StringBuilder b = new StringBuilder();
            b.Append('[');
            for (int i = 0; ; i++) {
                b.Append(a[i]);
                if (i == iMax)
                    return b.Append(']').ToString();
                b.Append(", ");
            }
        }

        /**
         * Returns a string representation of the contents of the specified array.
         * The string representation consists of a list of the array's elements,
         * enclosed in square brackets (<tt>"[]"</tt>).  Adjacent elements are
         * Separated by the characters <tt>", "</tt> (a comma followed by a
         * space).  Elements are Converted to strings as by
         * <tt>String.ValueOf(int)</tt>.  Returns <tt>"null"</tt> if <tt>a</tt> is
         * <tt>null</tt>.
         *
         * @param a the array whose string representation to return
         * @return a string representation of <tt>a</tt>
         * @since 1.5
         */
        public static String ToString(int[] a) {
            if (a == null)
                return "null";
            int iMax = a.Length - 1;
            if (iMax == -1)
                return "[]";

            StringBuilder b = new StringBuilder();
            b.Append('[');
            for (int i = 0; ; i++) {
                b.Append(a[i]);
                if (i == iMax)
                    return b.Append(']').ToString();
                b.Append(", ");
            }
        }

        /**
         * Returns a string representation of the contents of the specified array.
         * The string representation consists of a list of the array's elements,
         * enclosed in square brackets (<tt>"[]"</tt>).  Adjacent elements are
         * Separated by the characters <tt>", "</tt> (a comma followed by a
         * space).  Elements are Converted to strings as by
         * <tt>String.ValueOf(short)</tt>.  Returns <tt>"null"</tt> if <tt>a</tt>
         * is <tt>null</tt>.
         *
         * @param a the array whose string representation to return
         * @return a string representation of <tt>a</tt>
         * @since 1.5
         */
        public static String ToString(short[] a) {
            if (a == null)
                return "null";
            int iMax = a.Length - 1;
            if (iMax == -1)
                return "[]";

            StringBuilder b = new StringBuilder();
            b.Append('[');
            for (int i = 0; ; i++) {
                b.Append(a[i]);
                if (i == iMax)
                    return b.Append(']').ToString();
                b.Append(", ");
            }
        }

        /**
         * Returns a string representation of the contents of the specified array.
         * The string representation consists of a list of the array's elements,
         * enclosed in square brackets (<tt>"[]"</tt>).  Adjacent elements are
         * Separated by the characters <tt>", "</tt> (a comma followed by a
         * space).  Elements are Converted to strings as by
         * <tt>String.ValueOf(char)</tt>.  Returns <tt>"null"</tt> if <tt>a</tt>
         * is <tt>null</tt>.
         *
         * @param a the array whose string representation to return
         * @return a string representation of <tt>a</tt>
         * @since 1.5
         */
        public static String ToString(char[] a) {
            if (a == null)
                return "null";
            int iMax = a.Length - 1;
            if (iMax == -1)
                return "[]";

            StringBuilder b = new StringBuilder();
            b.Append('[');
            for (int i = 0; ; i++) {
                b.Append(a[i]);
                if (i == iMax)
                    return b.Append(']').ToString();
                b.Append(", ");
            }
        }

        /**
         * Returns a string representation of the contents of the specified array.
         * The string representation consists of a list of the array's elements,
         * enclosed in square brackets (<tt>"[]"</tt>).  Adjacent elements
         * are Separated by the characters <tt>", "</tt> (a comma followed
         * by a space).  Elements are Converted to strings as by
         * <tt>String.ValueOf(byte)</tt>.  Returns <tt>"null"</tt> if
         * <tt>a</tt> is <tt>null</tt>.
         *
         * @param a the array whose string representation to return
         * @return a string representation of <tt>a</tt>
         * @since 1.5
         */
        public static String ToString(byte[] a) {
            if (a == null)
                return "null";
            int iMax = a.Length - 1;
            if (iMax == -1)
                return "[]";

            StringBuilder b = new StringBuilder();
            b.Append('[');
            for (int i = 0; ; i++) {
                b.Append(a[i]);
                if (i == iMax)
                    return b.Append(']').ToString();
                b.Append(", ");
            }
        }

        /**
         * Returns a string representation of the contents of the specified array.
         * The string representation consists of a list of the array's elements,
         * enclosed in square brackets (<tt>"[]"</tt>).  Adjacent elements are
         * Separated by the characters <tt>", "</tt> (a comma followed by a
         * space).  Elements are Converted to strings as by
         * <tt>String.ValueOf(bool)</tt>.  Returns <tt>"null"</tt> if
         * <tt>a</tt> is <tt>null</tt>.
         *
         * @param a the array whose string representation to return
         * @return a string representation of <tt>a</tt>
         * @since 1.5
         */
        public static String ToString(bool[] a) {
            if (a == null)
                return "null";
            int iMax = a.Length - 1;
            if (iMax == -1)
                return "[]";

            StringBuilder b = new StringBuilder();
            b.Append('[');
            for (int i = 0; ; i++) {
                b.Append(a[i]);
                if (i == iMax)
                    return b.Append(']').ToString();
                b.Append(", ");
            }
        }

        /**
         * Returns a string representation of the contents of the specified array.
         * The string representation consists of a list of the array's elements,
         * enclosed in square brackets (<tt>"[]"</tt>).  Adjacent elements are
         * Separated by the characters <tt>", "</tt> (a comma followed by a
         * space).  Elements are Converted to strings as by
         * <tt>String.ValueOf(float)</tt>.  Returns <tt>"null"</tt> if <tt>a</tt>
         * is <tt>null</tt>.
         *
         * @param a the array whose string representation to return
         * @return a string representation of <tt>a</tt>
         * @since 1.5
         */
        public static String ToString(float[] a) {
            if (a == null)
                return "null";

            int iMax = a.Length - 1;
            if (iMax == -1)
                return "[]";

            StringBuilder b = new StringBuilder();
            b.Append('[');
            for (int i = 0; ; i++) {
                b.Append(a[i]);
                if (i == iMax)
                    return b.Append(']').ToString();
                b.Append(", ");
            }
        }

        /**
         * Returns a string representation of the contents of the specified array.
         * The string representation consists of a list of the array's elements,
         * enclosed in square brackets (<tt>"[]"</tt>).  Adjacent elements are
         * Separated by the characters <tt>", "</tt> (a comma followed by a
         * space).  Elements are Converted to strings as by
         * <tt>String.ValueOf(double)</tt>.  Returns <tt>"null"</tt> if <tt>a</tt>
         * is <tt>null</tt>.
         *
         * @param a the array whose string representation to return
         * @return a string representation of <tt>a</tt>
         * @since 1.5
         */
        public static String ToString(double[] a) {
            if (a == null)
                return "null";
            int iMax = a.Length - 1;
            if (iMax == -1)
                return "[]";

            StringBuilder b = new StringBuilder();
            b.Append('[');
            for (int i = 0; ; i++) {
                b.Append(a[i]);
                if (i == iMax)
                    return b.Append(']').ToString();
                b.Append(", ");
            }
        }


        /**
         * Returns a string representation of the "deep contents" of the specified
         * array.  If the array Contains other arrays as elements, the string
         * representation Contains their contents and so on.  This method is
         * designed for Converting multidimensional arrays to strings.
         *
         * The string representation consists of a list of the array's
         * elements, enclosed in square brackets (<tt>"[]"</tt>).  Adjacent
         * elements are Separated by the characters <tt>", "</tt> (a comma
         * followed by a space).  Elements are Converted to strings as by
         * <tt>String.ValueOf(Object)</tt>, unless they are themselves
         * arrays.
         *
         * If an element <tt>e</tt> is an array of a primitive type, it is
         * Converted to a string as by invoking the appropriate overloading of
         * <tt>Arrays.ToString(e)</tt>.  If an element <tt>e</tt> is an array of a
         * reference type, it is Converted to a string as by invoking
         * this method recursively.
         *
         * To avoid infInite recursion, if the specified array Contains itself
         * as an element, or Contains an indirect reference to itself through one
         * or more levels of arrays, the self-reference is Converted to the string
         * <tt>"[...]"</tt>.  For example, an array Containing only a reference
         * to itself would be rendered as <tt>"[[...]]"</tt>.
         *
         * This method returns <tt>"null"</tt> if the specified array
         * is <tt>null</tt>.
         *
         * @param a the array whose string representation to return
         * @return a string representation of <tt>a</tt>
         * @see #ToString(Object[])
         * @since 1.5
         */
        public static String DeepToString(Object[] a) {
            if (a == null)
                return "null";

            int bufLen = 20 * a.Length;
            if (a.Length != 0 && bufLen <= 0)
                bufLen = Int32.MaxValue;
            StringBuilder buf = new StringBuilder(bufLen);
            DeepToString(a, buf, new NPOI.Util.Collections.HashSet<Object[]>());
            return buf.ToString();
        }

        private static void DeepToString(Object[] a, StringBuilder buf,
                                         NPOI.Util.Collections.HashSet<Object[]> dejaVu)
        {
            if (a == null)
            {
                buf.Append("null");
                return;
            }
            int iMax = a.Length - 1;
            if (iMax == -1)
            {
                buf.Append("[]");
                return;
            }

            dejaVu.Add(a);
            buf.Append('[');
            for (int i = 0; ; i++)
            {

                Object element = a[i];
                if (element == null)
                {
                    buf.Append("null");
                }
                else
                {
                    Type eClass = element.GetType();
                    //Class<?> eClass = element.Class;

                    if (eClass.IsArray)
                    {
                        if (eClass == typeof(byte[]))
                            buf.Append(ToString((byte[])element));
                        else if (eClass == typeof(short[]))
                            buf.Append(ToString((short[])element));
                        else if (eClass == typeof(int[]))
                            buf.Append(ToString((int[])element));
                        else if (eClass == typeof(long[]))
                            buf.Append(ToString((long[])element));
                        else if (eClass == typeof(char[]))
                            buf.Append(ToString((char[])element));
                        else if (eClass == typeof(float[]))
                            buf.Append(ToString((float[])element));
                        else if (eClass == typeof(double[]))
                            buf.Append(ToString((double[])element));
                        else if (eClass == typeof(bool[]))
                            buf.Append(ToString((bool[])element));
                        else
                        { // element is an array of object references
                            if (dejaVu.Contains((element as object[])))
                                buf.Append("[...]");
                            else
                                DeepToString((Object[])element, buf, dejaVu);
                        }
                    }
                    else
                    {  // element is non-null and not an array
                        buf.Append(element.ToString());
                    }
                }
                if (i == iMax)
                    break;
                buf.Append(", ");
            }
            buf.Append(']');
            dejaVu.Remove(a);
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
