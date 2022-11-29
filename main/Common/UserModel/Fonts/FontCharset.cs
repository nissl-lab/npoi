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

namespace NPOI.Common.UserModel.Fonts
{
	using System;
	using System.Collections.Generic;
	using System.Text;

	/**
	 * Charset represents the basic set of characters associated with a font (that it can display), and
	 * corresponds to the ANSI codepage (8-bit or DBCS) of that character set used by a given language.
	 *
	 * @since POI 3.17-beta2
	 */
	public enum FontCharset
	{
		/** Specifies the English character set. */
		ANSI = 0x00000000,
		/**
	     * Specifies a character set based on the current system locale;
	     * for example, when the system locale is United States English,
	     * the default character set is ANSI_CHARSET.
	     */
		DEFAULT = 0x00000001, 
		/** Specifies a character set of symbols. */
		SYMBOL = 0x00000002,
		/** Specifies the Apple Macintosh character set. */
		MAC = 0x0000004D, 
		/** Specifies the Japanese character set. */
		SHIFTJIS = 0x00000080,
		/** Also spelled "Hangeul". Specifies the Hangul Korean character set. */
		HANGUL = 0x00000081, 
		/** Also spelled "Johap". Specifies the Johab Korean character set. */
		JOHAB = 0x00000082,
		/** Specifies the "simplified" Chinese character set for People's Republic of China. */
		GB2312 = 0x00000086, 
		/**
	     * Specifies the "traditional" Chinese character set, used mostly in
	     * Taiwan and in the Hong Kong and Macao Special Administrative Regions.
	     */
		CHINESEBIG5 = 0x00000088,
		/** Specifies the Greek character set. */
		GREEK = 0x000000A1,
		/** Specifies the Turkish character set. */
		TURKISH = 0x000000A2,
		/** Specifies the Vietnamese character set. */
		VIETNAMESE = 0x000000A3,
		/** Specifies the Hebrew character set. */
		HEBREW = 0x000000B1, 
		/** Specifies the Arabic character set. */
		ARABIC = 0x000000B2,
		/** Specifies the Baltic (Northeastern European) character set. */
		BALTIC = 0x000000BA,
		/** Specifies the Russian Cyrillic character set. */
		RUSSIAN = 0x000000CC,
		/** Specifies the Thai character set. */
		THAI = 0x000000DE,
		/** Specifies a Eastern European character set. */
		EASTEUROPE = 0x000000EE, 
		/**
	     * Specifies a mapping to one of the OEM code pages,
	     * according to the current system locale setting.
	     */
		OEM = 0x000000FF
	}
}
