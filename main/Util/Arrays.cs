
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
    }
}
