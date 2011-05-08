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

namespace NPOI.SS.UserModel
{

    /**
     *  The enumeration value indicating the possible paper size for a sheet
     *
     * @author Daniele Montagni
     */
    public enum PaperSize : short
    {
        US_Letter = 1,
        US_Letter_Small = 2,
        US_Tabloid = 3,
        US_Ledger = 4,
        US_Legal = 5,
        US_Statement = 6,
        US_Executive = 7,
        A3 = 8,
        A4 = 9,
        A4_Small = 10,
        A5 = 11,
        B4 = 12,
        B5 = 13,
        Folio = 14,
        Quarto = 15,
        TEN_BY_FOURTEEN = 16,
        ELEVEN_BY_SEVENTEEN = 17,
        US_Note = 18,
        US_Envelope_9 = 19,
        US_Envelope_10 = 20,
        US_Envelope_11 = 21,
        US_Envelope_12 = 22,
        US_Envelope_14 = 23,
        C_Size_Sheet = 24,
        D_Size_Sheet = 25,
        E_Size_Sheet = 26,
        Envelope_DL = 27,
        Envelope_C5 = 28,
        Envelope_C3 = 29,
        Envelope_C4 = 30,
        Envelope_C6 = 31,
        Envelope_MONARCH = 37,
        A4_EXTRA = 53,
        /// <summary>
        /// A4 Transverse - 210x297 mm 
        /// </summary>
        A4_TRANSVERSE_PAPERSIZE = 55,
        /// <summary>
        /// A4 Plus - 210x330 mm 
        /// </summary>
        A4_PLUS_PAPERSIZE = 60,
        /// <summary>
        /// US Letter Rotated 11 x 8 1/2 in
        /// </summary>
        LETTER_ROTATED_PAPERSIZE = 75,
        /// <summary>
        /// A4 Rotated - 297x210 mm */
        /// </summary>
        A4_ROTATED_PAPERSIZE = 77
    }
}