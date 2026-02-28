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

namespace NPOI.HWPF.Model
{

    public interface CharIndexTranslator
    {
        /**
 * Calculates the byte index of the given char index.
 * 
 * @param charPos
 *            The char position
 * @return The byte index
 */
        int GetByteIndex(int charPos);
        /**
         * Calculates the char index of the given byte index.
         * Look forward if index is not in table
         *
         * @param bytePos The character offset to check 
         * @return the char index
         */
        int GetCharIndex(int bytePos);

        /**
         * Calculates the char index of the given byte index.
         * Look forward if index is not in table
         *
         * @param bytePos The character offset to check
         * @param startCP look from this characted position 
         * @return the char index
         */
        int GetCharIndex(int bytePos, int startCP);
        /**
 * Finds character ranges that includes specified byte range.
 * 
 * @param startBytePosInclusive
 *            start byte range
 * @param endBytePosExclusive
 *            end byte range
 */
//        int[][] GetCharIndexRanges(int startBytePosInclusive,   int endBytePosExclusive);

        /**
         * Check if index is in table
         *
         * @param bytePos
         * @return true if index in table, false if not
         */

        bool IsIndexInTable(int bytePos);

        /**
         * Return first index >= bytePos that is in table
         *
         * @param bytePos
         * @return  first index greater or equal to bytePos that is in table
         */
        int LookIndexForward(int bytePos);

        /**
         * Return last index <= bytePos that is in table
         *
         * @param bytePos
         * @return last index less of equal to bytePos that is in table
         */
        int LookIndexBackward(int bytePos);

    }


}