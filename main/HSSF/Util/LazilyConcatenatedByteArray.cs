/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.HSSF.Util
{
    using System;
    using System.Collections.Generic;

    /**
     * Utility for delaying the concatenation of multiple byte arrays.  Doing this up-front
     * causes significantly more copying, which for a large number of byte arrays can cost
     * a large amount of time.
     */
    public class LazilyConcatenatedByteArray
    {
        private List<byte[]> arrays = new List<byte[]>(1);

        /**
         * Clears the array (sets the concatenated length back to zero.
         */
        public void Clear()
        {
            arrays.Clear();
        }

        /**
         * Concatenates an array onto the end of our array.
         * This is a relatively fast operation.
         *
         * @param array the array to concatenate.
         * @throws ArgumentException if {@code array} is {@code null}.
         */
        public void Concatenate(byte[] array)
        {
            if (array == null)
            {
                throw new ArgumentException("array cannot be null");
            }
            arrays.Add(array);
        }

        /**
         * Gets the concatenated contents as a single byte array.
         *
         * This is a slower operation, but the concatenated array is stored off as a single
         * array again so that subsequent calls will not perform Additional copying.
         *
         * @return the byte array.  Returns {@code null} if no data has been placed into it.
         */
        public byte[] ToArray()
        {
            if (arrays.Count==0)
            {
                return null;
            }
            else if (arrays.Count > 1)
            {
                int totalLength = 0;
                foreach (byte[] array in arrays)
                {
                    totalLength += array.Length;
                }

                byte[] concatenated = new byte[totalLength];
                int destPos = 0;
                foreach (byte[] array in arrays)
                {
                    Array.Copy(array, 0, concatenated, destPos, array.Length);
                    destPos += array.Length;
                }

                arrays.Clear();
                arrays.Add(concatenated);
            }

            return arrays[(0)];
        }
    }


}