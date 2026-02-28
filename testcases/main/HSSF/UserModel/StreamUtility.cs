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
namespace TestCases.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.UserModel;
    using System.IO;
    using System.Collections;

    /**
     * Utility class to help Test code verify that generated files do not differ from proof copies in 
     * any significant detail.  Normally this task would be simple except for the presence of artifacts
     * in the file that change every time it is generated.  Usually these volatile artifacts are  
     * time-stamps, user names, or other machine dependent parameters.
     *  
     * @author Josh Micich
     */
    public class StreamUtility
    {

        /**
         * Compares two streams with expected differences in specified regions.  The streams are
         * expected to be of equal Length and comparison is always byte for byte.  That is -
         * differences can only involve exchanging each individual byte for another single byte.<br/>
         * Both input streams are Closed.
         *  
         * @param allowableDifferenceRegions array of integer pairs: (offset, Length). 
         * Any differences encountered in these regions of the streams will be ignored
         * @return <c>null</c> if streams are identical, else the 
         * byte indexes of differing data.  If streams were different Lengths,
         * the returned indexes will be -1 and the Length of the shorter stream
         */
        public static int[] DiffStreams(Stream isA, Stream isB, int[] allowableDifferenceRegions)
        {

            if ((allowableDifferenceRegions.Length % 2) != 0)
            {
                throw new ArgumentException("allowableDifferenceRegions Length is odd");
            }
            bool success = false;
            int[] result;
            try
            {
                result = DiffInternal(isA, isB, allowableDifferenceRegions);
                success = true;
            }
            catch (IOException)
            {
                throw;
            }
            finally
            {
                Close(isA, success);
                Close(isB, success);
            }
            return result;
        }

        /**
         * @param success <c>false</c> if the outer method is throwing an exception.
         */
        private static void Close(Stream is1, bool success)
        {
            try
            {
                is1.Close();
            }
            catch (IOException)
            {
                if (success)
                {
                    // this is a new error. ok to throw
                    throw;
                }
                // else don't subvert original exception. just print stack trace for this one
                //e.printStackTrace();
            }
        }

        private static int[] DiffInternal(Stream isA, Stream isB, int[] allowableDifferenceRegions)
        {
            int offset = 0;
            ArrayList temp = new ArrayList();
            while (true)
            {
                int b = isA.ReadByte();
                int b2 = isB.ReadByte();
                if (b == -1)
                {
                    // EOF
                    if (b2 == -1)
                    {
                        return ToPrimitiveIntArray(temp);
                    }
                    return new int[] { -1, offset, };
                }
                if (b2 == -1)
                {
                    return new int[] { -1, offset, };
                }
                if (b != b2 && !IsIgnoredRegion(allowableDifferenceRegions, offset))
                {
                    temp.Add(offset);
                }
                offset++;
            }
        }

        private static bool IsIgnoredRegion(int[] allowableDifferenceRegions, int offset)
        {
            for (int i = 0; i < allowableDifferenceRegions.Length; i += 2)
            {
                int start = allowableDifferenceRegions[i];
                int end = start + allowableDifferenceRegions[i + 1];
                if (start <= offset && offset < end)
                {
                    return true;
                }
            }
            return false;
        }

        private static int[] ToPrimitiveIntArray(ArrayList temp)
        {
            int nItems = temp.Count;
            if (nItems < 1)
            {
                return null;
            }
             
             int[] boxInts = (int[])temp.ToArray(typeof(int));

            int[] result = new int[nItems];
            for (int i = 0; i < result.Length; i++)
            {
                result[i] = boxInts[i];
            }
            return result;
        }
    }
}