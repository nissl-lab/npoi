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

using NPOI.Common.UserModel.Fonts;
using NPOI.HSLF.Exceptions;
using NPOI.Util;
using System;
using System.Collections.Generic;

namespace NPOI.HSLF.Record
{
	public class FontEntityAtom : RecordAtom
	{
		private static int[] FLAGS_MASKS = {
		0x0001, 0x0100, 0x0200, 0x0400, 0x0800
		};

		private static string[] FLAGS_NAMES = {
		"EMBED_SUBSETTED",
		"RASTER_FONT",
		"DEVICE_FONT",
		"TRUETYPE_FONT",
		"NO_FONT_SUBSTITUTION"
	};

		/**
		 * record header
		 */
		private byte[] _header;

		/**
		 * record data
		 */
		private byte[] _recdata;

		/**
		 * Build an instance of <code>FontEntityAtom</code> from on-disk data
		 */
		/* package */
		FontEntityAtom(byte[] source, int start, int len)
		{
			// Get the header
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Grab the record data
			_recdata = IOUtils.SafelyClone(source, start + 8, len - 8, GetMaxRecordLength());
		}

		/**
		 * Create a new instance of <code>FontEntityAtom</code>
		 */
		public FontEntityAtom()
		{
			_recdata = new byte[68];
			_header = new byte[8];

			LittleEndian.PutShort(_header, 2, (short)GetRecordType());
			LittleEndian.PutInt(_header, 4, _recdata.Length);
		}

		//@Override
		public override long GetRecordType()
		{
			return RecordTypes.FontEntityAtom.typeID;
		}

		/**
		 * A null-terminated string that specifies the typeface name of the font.
		 * The length of this string must not exceed 32 characters
		 *  including the null terminator.
		 * @return font name
		 */
		public string GetFontName()
		{
			int maxLen = Math.Min(_recdata.Length, 64) / 2;
			return StringUtil.GetFromUnicodeLE0Terminated(_recdata, 0, maxLen);
		}

		/**
		 * Set the name of the font.
		 * The length of this string must not exceed 32 characters
		 *  including the null terminator.
		 * Will be converted to null-terminated if not already
		 * @param name of the font
		 */
		public void SetFontName(String name)
		{
			// Ensure it's not now too long
			int nameLen = name.Length + (name.EndsWith("\u0000") ? 0 : 1);
			if (nameLen > 32)
			{
				throw new HSLFException("The length of the font name, including null termination, must not exceed 32 characters");
			}

			// Everything's happy, so save the name
			byte[] bytes = StringUtil.GetToUnicodeLE(name);
			Array.Copy(bytes, 0, _recdata, 0, bytes.Length);
			// null the remaining bytes
			Arrays.Fill(_recdata, bytes.Length, 64, (byte)0);
		}

		public void SetFontIndex(int idx)
		{
			LittleEndian.PutShort(_header, 0, (short)idx);
		}

		public int GetFontIndex()
		{
			return LittleEndian.GetShort(_header, 0) >> 4;
		}

		/**
		 * set the character set
		 *
		 * @param charset - characterset
		 */
		public void SetCharSet(int charset)
		{
			_recdata[64] = (byte)charset;
		}

		/**
		 * get the character set
		 *
		 * @return charset - characterset
		 */
		public int GetCharSet()
		{
			return _recdata[64];
		}

		/**
		 * set the font flags
		 * Bit 1: If set, font is subsetted
		 *
		 * @param flags - the font flags
		 */
		public void SetFontFlags(int flags)
		{
			_recdata[65] = (byte)flags;
		}

		/**
		 * get the character set
		 * Bit 1: If set, font is subsetted
		 *
		 * @return the font flags
		 */
		public int GetFontFlags()
		{
			return _recdata[65];
		}

		/**
		 * set the font type
		 * <p>
		 * Bit 1: Raster Font
		 * Bit 2: Device Font
		 * Bit 3: TrueType Font
		 * </p>
		 *
		 * @param type - the font type
		 */
		public void SetFontType(int type)
		{
			_recdata[66] = (byte)type;
		}

		/**
		 * get the font type
		 * <p>
		 * Bit 1: Raster Font
		 * Bit 2: Device Font
		 * Bit 3: TrueType Font
		 * </p>
		 *
		 * @return the font type
		 */
		public int GetFontType()
		{
			return _recdata[66];
		}

		/**
		 * set lfPitchAndFamily
		 *
		 *
		 * @param val - Corresponds to the lfPitchAndFamily field of the Win32 API LOGFONT structure
		 */
		public void SetPitchAndFamily(int val)
		{
			_recdata[67] = (byte)val;
		}

		/**
		 * get lfPitchAndFamily
		 *
		 * @return corresponds to the lfPitchAndFamily field of the Win32 API LOGFONT structure
		 */
		public int GetPitchAndFamily()
		{
			return _recdata[67];
		}

		/**
		 * Write the contents of the record back, so it can be written to disk
		 */
		//@Override
		public override void WriteOut(OutputStream _out)
		{
			_out.Write(_header);
			_out.Write(_recdata);
		}

		//@Override
		public override IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
				"fontName", () => GetFontName(),
				"fontIndex", () => GetFontIndex(),
				"charset", () => GetCharSet(),
				"fontFlags", () => GenericRecordUtil.GetBitsAsString(GetFontFlags, FLAGS_MASKS, FLAGS_NAMES),
				"fontPitch", () => FontPitch.ValueOfPitchFamily((byte)GetPitchAndFamily()),
				"fontFamily", () => FontFamily.ValueOfPitchFamily((byte)GetPitchAndFamily())
			);
		}
	}
}