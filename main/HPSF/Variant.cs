/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS Is" BASIS,
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

namespace NPOI.HPSF
{
    using System;
    using System.Collections;


    /// <summary>
    /// The <em>Variant</em> types as defined by Microsoft's COM. I
    /// found this information in <a href="http://www.marin.clara.net/COM/variant_type_definitions.htm">
    /// http://www.marin.clara.net/COM/variant_type_definitions.htm</a>.
    /// In the variant types descriptions the following shortcuts are
    /// used: <strong> [V]</strong> - may appear in a VARIANT,
    /// <strong>[T]</strong> - may appear in a TYPEDESC,
    /// <strong>[P]</strong> - may appear in an OLE property Set,
    /// <strong>[S]</strong> - may appear in a Safe Array.
    /// @author Rainer Klute (klute@rainer-klute.de)
    /// @since 2002-02-09
    /// </summary>
    public class Variant
    {

        /**
         * [V][P] Nothing, i.e. not a single byte of data.
         */
        public const int VT_EMPTY = 0;

        /**
         * [V][P] SQL style Null.
         */
        public const int VT_NULL = 1;

        /**
         * [V][T][P][S] 2 byte signed int.
         */
        public const int VT_I2 = 2;

        /**
         * [V][T][P][S] 4 byte signed int.
         */
        public const int VT_I4 = 3;

        /**
         * [V][T][P][S] 4 byte real.
         */
        public const int VT_R4 = 4;

        /**
         * [V][T][P][S] 8 byte real.
         */
        public const int VT_R8 = 5;

        /**
         * [V][T][P][S] currency. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_CY = 6;

        /**
         * [V][T][P][S] DateTime. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_DATE = 7;

        /**
         * [V][T][P][S] OLE Automation string. <span
         * style="background-color: #ffff00">How long is this? How is it
         * To be interpreted?</span>
         */
        public const int VT_BSTR = 8;

        /**
         * [V][T][P][S] IDispatch *. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_DISPATCH = 9;

        /**
         * [V][T][S] SCODE. <span style="background-color: #ffff00">How
         * long is this? How is it To be interpreted?</span>
         */
        public const int VT_ERROR = 10;

        /**
         * [V][T][P][S] True=-1, False=0.
         */
        public const int VT_BOOL = 11;

        /**
         * [V][T][P][S] VARIANT *. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_VARIANT = 12;

        /**
         * [V][T][S] IUnknown *. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_UNKNOWN = 13;

        /**
         * [V][T][S] 16 byte fixed point.
         */
        public const int VT_DECIMAL = 14;

        /**
         * [T] signed char.
         */
        public const int VT_I1 = 16;

        /**
         * [V][T][P][S] unsigned char.
         */
        public const int VT_UI1 = 17;

        /**
         * [T][P] unsigned short.
         */
        public const int VT_UI2 = 18;

        /**
         * [T][P] unsigned int.
         */
        public const int VT_UI4 = 19;

        /**
         * [T][P] signed 64-bit int.
         */
        public const int VT_I8 = 20;

        /**
         * [T][P] unsigned 64-bit int.
         */
        public const int VT_UI8 = 21;

        /**
         * [T] signed machine int.
         */
        public const int VT_INT = 22;

        /**
         * [T] unsigned machine int.
         */
        public const int VT_UINT = 23;

        /**
         * [T] C style void.
         */
        public const int VT_VOID = 24;

        /**
         * [T] Standard return type. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_HRESULT = 25;

        /**
         * [T] pointer type. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_PTR = 26;

        /**
         * [T] (use VT_ARRAY in VARIANT).
         */
        public const int VT_SAFEARRAY = 27;

        /**
         * [T] C style array. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_CARRAY = 28;

        /**
         * [T] user defined type. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_USERDEFINED = 29;

        /**
         * [T][P] null terminated string.
         */
        public const int VT_LPSTR = 30;

        /**
         * [T][P] wide (Unicode) null terminated string.
         */
        public const int VT_LPWSTR = 31;

        /**
         * [P] FILETIME. The FILETIME structure holds a DateTime and time
         * associated with a file. The structure identifies a 64-bit
         * integer specifying the number of 100-nanosecond intervals which
         * have passed since January 1, 1601. This 64-bit value is split
         * into the two dwords stored in the structure.
         */
        public const int VT_FILETIME = 64;

        /**
         * [P] Length prefixed bytes.
         */
        public const int VT_BLOB = 65;

        /**
         * [P] Name of the stream follows.
         */
        public const int VT_STREAM = 66;

        /**
         * [P] Name of the storage follows.
         */
        public const int VT_STORAGE = 67;

        /**
         * [P] Stream Contains an object. <span
         * style="background-color: #ffff00"> How long is this? How is it
         * To be interpreted?</span>
         */
        public const int VT_STREAMED_OBJECT = 68;

        /**
         * [P] Storage Contains an object. <span
         * style="background-color: #ffff00"> How long is this? How is it
         * To be interpreted?</span>
         */
        public const int VT_STORED_OBJECT = 69;

        /**
         * [P] Blob Contains an object. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_BLOB_OBJECT = 70;

        /**
         * [P] Clipboard format. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_CF = 71;

        /**
         * [P] A Class ID.
         *
         * It consists of a 32 bit unsigned integer indicating the size
         * of the structure, a 32 bit signed integer indicating (Clipboard
         * Format Tag) indicating the type of data that it Contains, and
         * then a byte array containing the data.
         *
         * The valid Clipboard Format Tags are:
         *
         * <ul>
         *  <li>{@link Thumbnail#CFTAG_WINDOWS}</li>
         *  <li>{@link Thumbnail#CFTAG_MACINTOSH}</li>
         *  <li>{@link Thumbnail#CFTAG_NODATA}</li>
         *  <li>{@link Thumbnail#CFTAG_FMTID}</li>
         * </ul>
         *
         * <pre>typedef struct tagCLIPDATA {
         * // cbSize is the size of the buffer pointed To
         * // by pClipData, plus sizeof(ulClipFmt)
         * ULONG              cbSize;
         * long               ulClipFmt;
         * BYTE*              pClipData;
         * } CLIPDATA;</pre>
         *
         * See <a
         * href="msdn.microsoft.com/library/en-us/com/stgrstrc_0uwk.asp"
         * tarGet="_blank">
         * msdn.microsoft.com/library/en-us/com/stgrstrc_0uwk.asp</a>.
         */
        public const int VT_CLSID = 72;
        /**
         * "MUST be a VersionedStream. The storage representing the (non-simple)
         * property set MUST have a stream element with the name in the StreamName
         * field." -- [MS-OLEPS] -- v20110920; Object Linking and Embedding (OLE)
         * Property Set Data Structures; page 24 / 63
         */
        public const int VT_VERSIONED_STREAM = 0x0049;
        /**
         * [P] simple counted array. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_VECTOR = 0x1000;

        /**
         * [V] SAFEARRAY*. <span style="background-color: #ffff00">How
         * long is this? How is it To be interpreted?</span>
         */
        public const int VT_ARRAY = 0x2000;

        /**
         * [V] void* for local use. <span style="background-color:
         * #ffff00">How long is this? How is it To be
         * interpreted?</span>
         */
        public const int VT_BYREF = 0x4000;

        /**
         * FIXME (3): Document this!
         */
        public const int VT_RESERVED = 0x8000;

        /**
         * FIXME (3): Document this!
         */
        public const int VT_ILLEGAL = 0xFFFF;

        /**
         * FIXME (3): Document this!
         */
        public const int VT_ILLEGALMASKED = 0xFFF;

        /**
         * FIXME (3): Document this!
         */
        public const int VT_TYPEMASK = 0xFFF;



        /**
         * Maps the numbers denoting the variant types To their corresponding
         * variant type names.
         */
        private static IDictionary numberToName;

        private static IDictionary numberToLength;

        /**
         * Denotes a variant type with a Length that is unknown To HPSF yet.
         */
        public const int Length_UNKNOWN = -2;

        /**
         * Denotes a variant type with a variable Length.
         */
        public const int Length_VARIABLE = -1;

        /**
         * Denotes a variant type with a Length of 0 bytes.
         */
        public const int Length_0 = 0;

        /**
         * Denotes a variant type with a Length of 2 bytes.
         */
        public const int Length_2 = 2;

        /**
         * Denotes a variant type with a Length of 4 bytes.
         */
        public const int Length_4 = 4;

        /**
         * Denotes a variant type with a Length of 8 bytes.
         */
        public const int Length_8 = 8;



        static Variant()
        {
            /* Initialize the number-to-name map: */
            Hashtable tm1 = new Hashtable();
            tm1[0] = "VT_EMPTY";
            tm1[1] = "VT_NULL";
            tm1[2] = "VT_I2";
            tm1[3] = "VT_I4";
            tm1[4] = "VT_R4";
            tm1[5] = "VT_R8";
            tm1[6] = "VT_CY";
            tm1[7] = "VT_DATE";
            tm1[8] = "VT_BSTR";
            tm1[9] = "VT_DISPATCH";
            tm1[10] = "VT_ERROR";
            tm1[11] = "VT_BOOL";
            tm1[12] = "VT_VARIANT";
            tm1[13] = "VT_UNKNOWN";
            tm1[14] = "VT_DECIMAL";
            tm1[16] = "VT_I1";
            tm1[17] = "VT_UI1";
            tm1[18] = "VT_UI2";
            tm1[19] = "VT_UI4";
            tm1[20] = "VT_I8";
            tm1[21] = "VT_UI8";
            tm1[22] = "VT_INT";
            tm1[23] = "VT_UINT";
            tm1[24] = "VT_VOID";
            tm1[25] = "VT_HRESULT";
            tm1[26] = "VT_PTR";
            tm1[27] = "VT_SAFEARRAY";
            tm1[28] = "VT_CARRAY";
            tm1[29] = "VT_USERDEFINED";
            tm1[30] = "VT_LPSTR";
            tm1[31] = "VT_LPWSTR";
            tm1[64] = "VT_FILETIME";
            tm1[65] = "VT_BLOB";
            tm1[66] = "VT_STREAM";
            tm1[67] = "VT_STORAGE";
            tm1[68] = "VT_STREAMED_OBJECT";
            tm1[69] = "VT_STORED_OBJECT";
            tm1[70] = "VT_BLOB_OBJECT";
            tm1[71] = "VT_CF";
            tm1[72] = "VT_CLSID";

            numberToName = tm1;

            /* Initialize the number-to-Length map: */
            Hashtable tm2 = new Hashtable();
            tm2[0] = Length_0;
            tm2[1] = Length_UNKNOWN;
            tm2[2] = Length_2;
            tm2[3] = Length_4;
            tm2[4] = Length_4;
            tm2[5] = Length_8;
            tm2[6] = Length_UNKNOWN;
            tm2[7] = Length_UNKNOWN;
            tm2[8] = Length_UNKNOWN;
            tm2[9] = Length_UNKNOWN;
            tm2[10] = Length_UNKNOWN;
            tm2[11] = Length_UNKNOWN;
            tm2[12] = Length_UNKNOWN;
            tm2[13] = Length_UNKNOWN;
            tm2[14] = Length_UNKNOWN;
            tm2[16] = Length_UNKNOWN;
            tm2[17] = Length_UNKNOWN;
            tm2[18] = Length_UNKNOWN;
            tm2[19] = Length_UNKNOWN;
            tm2[20] = Length_UNKNOWN;
            tm2[21] = Length_UNKNOWN;
            tm2[22] = Length_UNKNOWN;
            tm2[23] = Length_UNKNOWN;
            tm2[24] = Length_UNKNOWN;
            tm2[25] = Length_UNKNOWN;
            tm2[26] = Length_UNKNOWN;
            tm2[27] = Length_UNKNOWN;
            tm2[28] = Length_UNKNOWN;
            tm2[29] = Length_UNKNOWN;
            tm2[30] = Length_VARIABLE;
            tm2[31] = Length_UNKNOWN;
            tm2[64] = Length_8;
            tm2[65] = Length_UNKNOWN;
            tm2[66] = Length_UNKNOWN;
            tm2[67] = Length_UNKNOWN;
            tm2[68] = Length_UNKNOWN;
            tm2[69] = Length_UNKNOWN;
            tm2[70] = Length_UNKNOWN;
            tm2[71] = Length_UNKNOWN;
            tm2[72] = Length_UNKNOWN;

            numberToLength = tm2;
        }



        /// <summary>
        /// Returns the variant type name associated with a variant type
        /// number.
        /// </summary>
        /// <param name="variantType">The variant type number.</param>
        /// <returns>The variant type name or the string "unknown variant type"</returns>
        public static String GetVariantName(long variantType)
        {
            String name = (String)numberToName[variantType];
            return name != null ? name : "unknown variant type";
        }

        /// <summary>
        /// Returns a variant type's Length.
        /// </summary>
        /// <param name="variantType">The variant type number.</param>
        /// <returns>The Length of the variant type's data in bytes. If the Length Is
        /// variable, i.e. the Length of a string, -1 is returned. If HPSF does not
        /// know the Length, -2 is returned. The latter usually indicates an
        /// unsupported variant type.</returns>
        public static int GetVariantLength(long variantType)
        {
            long key = (int)variantType;
            if (numberToLength.Contains(key))
                return -2;
            long Length = (long)numberToLength[key];
            return Convert.ToInt32(Length);
        }

    }
}