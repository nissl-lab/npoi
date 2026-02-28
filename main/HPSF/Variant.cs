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
    using System.Collections.Generic;


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
        private static Dictionary<long,string> numberToName;

        private static Dictionary<long,int> numberToLength;

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

        private static Object[][] NUMBER_TO_NAME_LIST = [
            [  0L, "VT_EMPTY",           Length_0 ],
            [  1L, "VT_NULL",            Length_UNKNOWN ],
            [  2L, "VT_I2",              Length_2 ],
            [  3L, "VT_I4",              Length_4 ],
            [  4L, "VT_R4",              Length_4 ],
            [  5L, "VT_R8",              Length_8 ],
            [  6L, "VT_CY",              Length_UNKNOWN ],
            [  7L, "VT_DATE",            Length_UNKNOWN ],
            [  8L, "VT_BSTR",            Length_UNKNOWN ],
            [  9L, "VT_DISPATCH",        Length_UNKNOWN ],
            [ 10L, "VT_ERROR",           Length_UNKNOWN ],
            [ 11L, "VT_BOOL",            Length_UNKNOWN ],
            [ 12L, "VT_VARIANT",         Length_UNKNOWN ],
            [ 13L, "VT_UNKNOWN",         Length_UNKNOWN ],
            [ 14L, "VT_DECIMAL",         Length_UNKNOWN ],
            [ 16L, "VT_I1",              Length_UNKNOWN ],
            [ 17L, "VT_UI1",             Length_UNKNOWN ],
            [ 18L, "VT_UI2",             Length_UNKNOWN ],
            [ 19L, "VT_UI4",             Length_UNKNOWN ],
            [ 20L, "VT_I8",              Length_UNKNOWN ],
            [ 21L, "VT_UI8",             Length_UNKNOWN ],
            [ 22L, "VT_INT",             Length_UNKNOWN ],
            [ 23L, "VT_UINT",            Length_UNKNOWN ],
            [ 24L, "VT_VOID",            Length_UNKNOWN ],
            [ 25L, "VT_HRESULT",         Length_UNKNOWN ],
            [ 26L, "VT_PTR",             Length_UNKNOWN ],
            [ 27L, "VT_SAFEARRAY",       Length_UNKNOWN ],
            [ 28L, "VT_CARRAY",          Length_UNKNOWN ],
            [ 29L, "VT_USERDEFINED",     Length_UNKNOWN ],
            [ 30L, "VT_LPSTR",           Length_VARIABLE ],
            [ 31L, "VT_LPWSTR",          Length_UNKNOWN ],
            [ 64L, "VT_FILETIME",        Length_8 ],
            [ 65L, "VT_BLOB",            Length_UNKNOWN ],
            [ 66L, "VT_STREAM",          Length_UNKNOWN ],
            [ 67L, "VT_STORAGE",         Length_UNKNOWN ],
            [ 68L, "VT_STREAMED_OBJECT", Length_UNKNOWN ],
            [ 69L, "VT_STORED_OBJECT",   Length_UNKNOWN ],
            [ 70L, "VT_BLOB_OBJECT",     Length_UNKNOWN ],
            [ 71L, "VT_CF",              Length_UNKNOWN ],
            [ 72L, "VT_CLSID",           Length_UNKNOWN ]
        ];

        /* Initialize the number-to-name and number-to-length map: */
        static Variant() 
        {
            Dictionary<long,String> number2Name = new Dictionary<long,String>(NUMBER_TO_NAME_LIST.Length);
            Dictionary<long,int> number2Len = new Dictionary<long, int>(NUMBER_TO_NAME_LIST.Length);

            foreach (Object[] nn in NUMBER_TO_NAME_LIST)
            {
                number2Name[(long)nn[0]] = (String)nn[1];
                number2Len[(long)nn[0]] = (int)nn[2];
            }
            numberToName = number2Name;
            numberToLength = number2Len;
        }


        /// <summary>
        /// Returns the variant type name associated with a variant type
        /// number.
        /// </summary>
        /// <param name="variantType">The variant type number.</param>
        /// <returns>The variant type name or the string "unknown variant type"</returns>
        public static String GetVariantName(long variantType)
        {
            long vt = variantType;
            String name = "";
            if ((vt & VT_VECTOR) != 0)
            {
                name = "Vector of ";
                vt -= VT_VECTOR;
            }
            else if ((vt & VT_ARRAY) != 0)
            {
                name = "Array of ";
                vt -= VT_ARRAY;
            }
            else if ((vt & VT_BYREF) != 0)
            {
                name = "ByRef of ";
                vt -= VT_BYREF;
            }
            numberToName.TryGetValue(vt, out string name2);
            name += name2;
            return (name != null && !"".Equals(name)) ? name : "unknown variant type";
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
            if(numberToLength.TryGetValue(variantType, out int length))
            {
                return length;
            }
            else
            {
                return -2;
            }
        }

    }
}