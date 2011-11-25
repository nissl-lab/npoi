
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
    }
}
