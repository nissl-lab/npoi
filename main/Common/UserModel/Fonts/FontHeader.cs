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
    using NPOI.HSLF.Record;
    using NPOI.Util;
	using System;
	using System.Collections.Generic;
	using System.Text;

	/**
	 * The header data of an EOT font.<p>
	 *
	 * Currently only version 1 fields are read to identify a stream to be embedded.
	 *
	 * @see <a href="http://www.w3.org/Submission/EOT">Embedded OpenType (EOT) File Format</a>
	 */
	public class FontHeader : FontInfo, GenericRecord
	{
		public enum PanoseFamily
		{
			ANY, NO_FIT, TEXT_DISPLAY, SCRIPT, DECORATIVE, PICTORIAL
		}

		public enum PanoseSerif
		{
			ANY, NO_FIT, COVE, OBTUSE_COVE, SQUARE_COVE, OBTUSE_SQUARE_COVE, SQUARE, THIN, BONE,
			EXAGGERATED, TRIANGLE, NORMAL_SANS, OBTUSE_SANS, PERP_SANS, FLARED, ROUNDED
		}

		public enum PanoseWeight
		{
			ANY, NO_FIT, VERY_LIGHT, LIGHT, THIN, BOOK, MEDIUM, DEMI, BOLD, HEAVY, BLACK, NORD
		}

		public enum PanoseProportion
		{
			ANY, NO_FIT, OLD_STYLE, MODERN, EVEN_WIDTH, EXPANDED, CONDENSED, VERY_EXPANDED, VERY_CONDENSED, MONOSPACED
		}

		public enum PanoseContrast
		{
			ANY, NO_FIT, NONE, VERY_LOW, LOW, MEDIUM_LOW, MEDIUM, MEDIUM_HIGH, HIGH, VERY_HIGH
		}

		public enum PanoseStroke
		{
			ANY, NO_FIT, GRADUAL_DIAG, GRADUAL_TRAN, GRADUAL_VERT, GRADUAL_HORZ, RAPID_VERT, RAPID_HORZ, INSTANT_VERT
		}

		public enum PanoseArmStyle
		{
			ANY, NO_FIT, STRAIGHT_ARMS_HORZ, STRAIGHT_ARMS_WEDGE, STRAIGHT_ARMS_VERT, STRAIGHT_ARMS_SINGLE_SERIF,
			STRAIGHT_ARMS_DOUBLE_SERIF, BENT_ARMS_HORZ, BENT_ARMS_WEDGE, BENT_ARMS_VERT, BENT_ARMS_SINGLE_SERIF,
			BENT_ARMS_DOUBLE_SERIF,
		}

		public enum PanoseLetterForm
		{
			ANY, NO_FIT, NORMAL_CONTACT, NORMAL_WEIGHTED, NORMAL_BOXED, NORMAL_FLATTENED, NORMAL_ROUNDED,
			NORMAL_OFF_CENTER, NORMAL_SQUARE, OBLIQUE_CONTACT, OBLIQUE_WEIGHTED, OBLIQUE_BOXED, OBLIQUE_FLATTENED,
			OBLIQUE_ROUNDED, OBLIQUE_OFF_CENTER, OBLIQUE_SQUARE
		}

		public enum PanoseMidLine
		{
			ANY, NO_FIT, STANDARD_TRIMMED, STANDARD_POINTED, STANDARD_SERIFED, HIGH_TRIMMED, HIGH_POINTED, HIGH_SERIFED,
			CONSTANT_TRIMMED, CONSTANT_POINTED, CONSTANT_SERIFED, LOW_TRIMMED, LOW_POINTED, LOW_SERIFED
		}

		public enum PanoseXHeight
		{
			ANY, NO_FIT, CONSTANT_SMALL, CONSTANT_STD, CONSTANT_LARGE, DUCKING_SMALL, DUCKING_STD, DUCKING_LARGE
		}

		private static int[] FLAGS_MASKS = {
			0x00000001, 0x00000004, 0x00000010, 0x00000020, 0x00000040, 0x00000080, 0x10000000
		};

		private static String[] FLAGS_NAMES = {
	        "SUBSET", "TTCOMPRESSED", "FAILIFVARIATIONSIMULATED", "EMBEDEUDC", "VALIDATIONTESTS", "WEBOBJECT", "XORENCRYPTDATA"
	    };
	
	    private static int[] FSTYPE_MASKS = {
			0x0000, 0x0002, 0x0004, 0x0008, 0x0100, 0x0200
		};

		private static String[] FSTYPE_NAMES = {
	        "INSTALLABLE_EMBEDDING",
	        "RESTRICTED_LICENSE_EMBEDDING",
	        "PREVIEW_PRINT_EMBEDDING",
	        "EDITABLE_EMBEDDING",
	        "NO_SUBSETTING",
	        "BITMAP_EMBEDDING_ONLY"
	    };
	
	    /**
	     * Fonts with a font weight of 400 are regarded as regular weighted.
	     * Higher font weights (up to 1000) are bold - lower weights are thin.
	     */
	    public static int REGULAR_WEIGHT = 400;

		private int eotSize;
		private int fontDataSize;
		private int version;
		private int flags;
		private byte[] panose = new byte[10];
		private byte charset;
		private byte italic;
		private int weight;
		private int fsType;
		private int magic;
		private int unicodeRange1;
		private int unicodeRange2;
		private int unicodeRange3;
		private int unicodeRange4;
		private int codePageRange1;
		private int codePageRange2;
		private int checkSumAdjustment;
		private String familyName;
		private String styleName;
		private String versionName;
		private String fullName;

		public void init(byte[] source, int offset, int length)
		{
			init(new LittleEndianByteArrayInputStream(source, offset, length));
		}

		public void init(ILittleEndianInput leis)
		{
			eotSize = leis.ReadInt();
			fontDataSize = leis.ReadInt();
			version = leis.ReadInt();
			if (version != 0x00010000 && version != 0x00020001 && version != 0x00020002)
			{
				throw new RuntimeException("not a EOT font data stream");
			}
			flags = leis.ReadInt();
			leis.ReadFully(panose);
			charset = (byte)leis.ReadByte();
			italic = (byte)leis.ReadByte();
			weight = leis.ReadInt();
			fsType = leis.ReadUShort();
			magic = leis.ReadUShort();
			if (magic != 0x504C)
			{
				throw new RuntimeException("not a EOT font data stream");
			}
			unicodeRange1 = leis.ReadInt();
			unicodeRange2 = leis.ReadInt();
			unicodeRange3 = leis.ReadInt();
			unicodeRange4 = leis.ReadInt();
			codePageRange1 = leis.ReadInt();
			codePageRange2 = leis.ReadInt();
			checkSumAdjustment = leis.ReadInt();
			int reserved1 = leis.ReadInt();
			int reserved2 = leis.ReadInt();
			int reserved3 = leis.ReadInt();
			int reserved4 = leis.ReadInt();
			familyName = readName(leis);
			styleName = readName(leis);
			versionName = readName(leis);
			fullName = readName(leis);

		}

		public InputStream bufferInit(InputStream fontStream)
		{
			LittleEndianInputStream inputStream = new LittleEndianInputStream(fontStream);
			inputStream.Mark(1000);
			init(inputStream);
			inputStream.Reset();
			return inputStream;
		}

	private String readName(ILittleEndianInput leis)
	{
		// padding
		leis.ReadShort();
		int nameSize = leis.ReadUShort();
		byte[] nameBuf = IOUtils.SafelyAllocate(nameSize, 1000);
		leis.ReadFully(nameBuf);
		// may be 0-terminated, just trim it away
		return Encoding.Unicode.GetString(nameBuf).Trim();
	}

	public bool isItalic()
	{
		return italic != 0;
	}

	public int GetWeight()
	{
		return weight;
	}

	public bool isBold()
	{
		return GetWeight() > REGULAR_WEIGHT;
	}

	public byte getCharsetByte()
	{
		return charset;
	}

	public FontCharset GetCharset()
	{
		return (FontCharset)getCharsetByte();
	}

	public FontPitch GetPitch()
	{
		return FontPitch.ValueOf((int)GetPitchEnum());
	}

	public FontPitchEnum GetPitchEnum()
	{
		switch (getPanoseFamily())
		{
			default:
			case PanoseFamily.ANY:
			case PanoseFamily.NO_FIT:
				return FontPitchEnum.VARIABLE;

			// Latin Text
			case PanoseFamily.TEXT_DISPLAY:
			// Latin Decorative
			case PanoseFamily.DECORATIVE:
				return (getPanoseProportion() == PanoseProportion.MONOSPACED) ? FontPitchEnum.FIXED : FontPitchEnum.VARIABLE;

			// Latin Hand Written
			case PanoseFamily.SCRIPT:
			// Latin Symbol
			case PanoseFamily.PICTORIAL:
				return (getPanoseProportion() == PanoseProportion.MODERN) ? FontPitchEnum.FIXED : FontPitchEnum.VARIABLE;
		}

	}

	public FontFamily GetFamily()
	{
		return FontFamily.ValueOf((int)GetFamilyEnum());
	}

	public FontFamilyEnum GetFamilyEnum()
	{
		switch (getPanoseFamily())
		{
			case PanoseFamily.ANY:
			case PanoseFamily.NO_FIT:
				return FontFamilyEnum.FF_DONTCARE;
			// Latin Text
			case PanoseFamily.TEXT_DISPLAY:
				switch (getPanoseSerif())
				{
					case PanoseSerif.TRIANGLE:
					case PanoseSerif.NORMAL_SANS:
					case PanoseSerif.OBTUSE_SANS:
					case PanoseSerif.PERP_SANS:
					case PanoseSerif.FLARED:
					case PanoseSerif.ROUNDED:
						return FontFamilyEnum.FF_SWISS;
					default:
						return FontFamilyEnum.FF_ROMAN;
				}
			// Latin Hand Written
			case PanoseFamily.SCRIPT:
				return FontFamilyEnum.FF_SCRIPT;
			// Latin Decorative
			default:
			case PanoseFamily.DECORATIVE:
				return FontFamilyEnum.FF_DECORATIVE;
			// Latin Symbol
			case PanoseFamily.PICTORIAL:
				return FontFamilyEnum.FF_MODERN;
		}
	}

	public String getFamilyName()
	{
		return familyName;
	}

	public String getStyleName()
	{
		return styleName;
	}

	public String getVersionName()
	{
		return versionName;
	}

	public string getFullName()
	{
		return fullName;
	}

	public byte[] GetPanose()
	{
		return panose;
	}

	public string GetTypeface()
	{
		return getFamilyName();
	}

	public int getFlags()
	{
		return flags;
	}

		public IReadOnlyDictionary<string, Func<object>> getGenericProperties()
	{
            Dictionary<string, Func<object>> m = new Dictionary<string, Func<object>>();
		m.Add("eotSize", () => eotSize);
		m.Add("fontDataSize", () => fontDataSize);
		m.Add("version", () => version);
		m.Add("flags", GenericRecordUtil.GetBitsAsString(this.getFlags, FLAGS_MASKS, FLAGS_NAMES));
		m.Add("panose.familyType", () => getPanoseFamily());
		m.Add("panose.serifType", () => getPanoseSerif());
		m.Add("panose.weight", () => getPanoseWeight());
		m.Add("panose.proportion", () => getPanoseProportion());
		m.Add("panose.contrast", () => getPanoseContrast());
		m.Add("panose.stroke", () => getPanoseStroke());
		m.Add("panose.armStyle", () => getPanoseArmStyle());
		m.Add("panose.letterForm", () => getPanoseLetterForm());
		m.Add("panose.midLine", () => getPanoseMidLine());
		m.Add("panose.xHeight", () => getPanoseXHeight());
		m.Add("charset", () => GetCharset());
		m.Add("italic", () => isItalic());
		m.Add("weight", () => GetWeight());
		m.Add("fsType", GenericRecordUtil.GetBitsAsString(() => fsType, FSTYPE_MASKS, FSTYPE_NAMES));
		m.Add("unicodeRange1", () => unicodeRange1);
		m.Add("unicodeRange2", () => unicodeRange2);
		m.Add("unicodeRange3", () => unicodeRange3);
		m.Add("unicodeRange4", () => unicodeRange4);
		m.Add("codePageRange1", () => codePageRange1);
		m.Add("codePageRange2", () => codePageRange2);
		m.Add("checkSumAdjustment", () => checkSumAdjustment);
		m.Add("familyName", this.getFamilyName);
		m.Add("styleName", this.getStyleName);
		m.Add("versionName", this.getVersionName);
		m.Add("fullName", this.getFullName);
		return m;
	}

	public PanoseFamily getPanoseFamily()
	{
		return GenericRecordUtil.SafeEnum((PanoseFamily[])Enum.GetValues(typeof(PanoseFamily)), () => panose[0], PanoseFamily.ANY)();
	}

	public PanoseSerif getPanoseSerif()
	{
		return GenericRecordUtil.SafeEnum((PanoseSerif[])Enum.GetValues(typeof(PanoseSerif)), () => panose[1], PanoseSerif.ANY)();
	}

	public PanoseWeight getPanoseWeight()
	{
		return GenericRecordUtil.SafeEnum((PanoseWeight[])Enum.GetValues(typeof(PanoseWeight)), () => panose[2], PanoseWeight.ANY)();
	}

	public PanoseProportion getPanoseProportion()
	{
		return GenericRecordUtil.SafeEnum((PanoseProportion[])Enum.GetValues(typeof(PanoseProportion)), () => panose[3], PanoseProportion.ANY)();
	}

	public PanoseContrast getPanoseContrast()
	{
		return GenericRecordUtil.SafeEnum((PanoseContrast[])Enum.GetValues(typeof(PanoseContrast)), () => panose[4], PanoseContrast.ANY)();
	}

	public PanoseStroke getPanoseStroke()
	{
		return GenericRecordUtil.SafeEnum((PanoseStroke[])Enum.GetValues(typeof(PanoseStroke)), () => panose[5], PanoseStroke.ANY)();
	}

	public PanoseArmStyle getPanoseArmStyle()
	{
		return GenericRecordUtil.SafeEnum((PanoseArmStyle[])Enum.GetValues(typeof(PanoseArmStyle)), () => panose[6], PanoseArmStyle.ANY)();
	}

	public PanoseLetterForm getPanoseLetterForm()
	{
		return GenericRecordUtil.SafeEnum((PanoseLetterForm[])Enum.GetValues(typeof(PanoseLetterForm)), () => panose[7], PanoseLetterForm.ANY)();
	}

	public PanoseMidLine getPanoseMidLine()
	{
		return GenericRecordUtil.SafeEnum((PanoseMidLine[])Enum.GetValues(typeof(PanoseMidLine)), () => panose[8], PanoseMidLine.ANY)();
	}

	public PanoseXHeight getPanoseXHeight()
	{
		return GenericRecordUtil.SafeEnum((PanoseXHeight[])Enum.GetValues(typeof(PanoseXHeight)), () => panose[9], PanoseXHeight.ANY)();
	}

        public int GetIndex()
        {
            throw new NotImplementedException();
        }

        public void SetIndex(int index)
        {
            throw new NotImplementedException();
        }

        public void SetTypeface(string typeface)
        {
            throw new NotImplementedException();
        }

        public void SetCharset(FontCharset charset)
        {
            throw new NotImplementedException();
        }

        public void SetFamily(FontFamily family)
        {
            throw new NotImplementedException();
        }

        public void SetPitch(FontPitch pitch)
        {
            throw new NotImplementedException();
        }

        public void SetPanose(byte[] panose)
        {
            throw new NotImplementedException();
        }

        public List<T> GetFacets<T>()
        {
            throw new NotImplementedException();
        }

        public RecordTypes GetGenericRecordType()
        {
            throw new NotImplementedException();
        }

        public IDictionary<string, Func<T>> GetGenericProperties<T>()
        {
            throw new NotImplementedException();
        }

        public IList<GenericRecord> GetGenericChildren()
        {
            throw new NotImplementedException();
        }
    }
}
