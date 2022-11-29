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

		private static final int[] FLAGS_MASKS = {
			0x00000001, 0x00000004, 0x00000010, 0x00000020, 0x00000040, 0x00000080, 0x10000000
		};

		private static final String[] FLAGS_NAMES = {
	        "SUBSET", "TTCOMPRESSED", "FAILIFVARIATIONSIMULATED", "EMBEDEUDC", "VALIDATIONTESTS", "WEBOBJECT", "XORENCRYPTDATA"
	    };
	
	    private static final int[] FSTYPE_MASKS = {
			0x0000, 0x0002, 0x0004, 0x0008, 0x0100, 0x0200
		};

		private static final String[] FSTYPE_NAMES = {
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
	    public static final int REGULAR_WEIGHT = 400;

		private int eotSize;
		private int fontDataSize;
		private int version;
		private int flags;
		private final byte[] panose = new byte[10];
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

		public void init(LittleEndianInput leis)
		{
			eotSize = leis.readInt();
			fontDataSize = leis.readInt();
			version = leis.readInt();
			if (version != 0x00010000 && version != 0x00020001 && version != 0x00020002)
			{
				throw new RuntimeException("not a EOT font data stream");
			}
			flags = leis.readInt();
			leis.readFully(panose);
			charset = leis.readByte();
			italic = leis.readByte();
			weight = leis.readInt();
			fsType = leis.readUShort();
			magic = leis.readUShort();
			if (magic != 0x504C)
			{
				throw new RuntimeException("not a EOT font data stream");
			}
			unicodeRange1 = leis.readInt();
			unicodeRange2 = leis.readInt();
			unicodeRange3 = leis.readInt();
			unicodeRange4 = leis.readInt();
			codePageRange1 = leis.readInt();
			codePageRange2 = leis.readInt();
			checkSumAdjustment = leis.readInt();
			int reserved1 = leis.readInt();
			int reserved2 = leis.readInt();
			int reserved3 = leis.readInt();
			int reserved4 = leis.readInt();
			familyName = readName(leis);
			styleName = readName(leis);
			versionName = readName(leis);
			fullName = readName(leis);

		}

		public InputStream bufferInit(InputStream fontStream) throws IOException
		{
			LittleEndianInputStream is = new LittleEndianInputStream(fontStream);
	        is.mark(1000);
		init(is);
	        is.reset();
	        return is;
	    }

	private String readName(LittleEndianInput leis)
	{
		// padding
		leis.readShort();
		int nameSize = leis.readUShort();
		byte[] nameBuf = IOUtils.safelyAllocate(nameSize, 1000);
		leis.readFully(nameBuf);
		// may be 0-terminated, just trim it away
		return new String(nameBuf, 0, nameSize, StandardCharsets.UTF_16LE).trim();
	}

	public boolean isItalic()
	{
		return italic != 0;
	}

	public int getWeight()
	{
		return weight;
	}

	public boolean isBold()
	{
		return getWeight() > REGULAR_WEIGHT;
	}

	public byte getCharsetByte()
	{
		return charset;
	}

	public FontCharset getCharset()
	{
		return FontCharset.valueOf(getCharsetByte());
	}

	public FontPitch getPitch()
	{
		switch (getPanoseFamily())
		{
			default:
			case ANY:
			case NO_FIT:
				return FontPitch.VARIABLE;

			// Latin Text
			case TEXT_DISPLAY:
			// Latin Decorative
			case DECORATIVE:
				return (getPanoseProportion() == PanoseProportion.MONOSPACED) ? FontPitch.FIXED : FontPitch.VARIABLE;

			// Latin Hand Written
			case SCRIPT:
			// Latin Symbol
			case PICTORIAL:
				return (getPanoseProportion() == PanoseProportion.MODERN) ? FontPitch.FIXED : FontPitch.VARIABLE;
		}

	}

	public FontFamily getFamily()
	{
		switch (getPanoseFamily())
		{
			case ANY:
			case NO_FIT:
				return FontFamily.FF_DONTCARE;
			// Latin Text
			case TEXT_DISPLAY:
				switch (getPanoseSerif())
				{
					case TRIANGLE:
					case NORMAL_SANS:
					case OBTUSE_SANS:
					case PERP_SANS:
					case FLARED:
					case ROUNDED:
						return FontFamily.FF_SWISS;
					default:
						return FontFamily.FF_ROMAN;
				}
			// Latin Hand Written
			case SCRIPT:
				return FontFamily.FF_SCRIPT;
			// Latin Decorative
			default:
			case DECORATIVE:
				return FontFamily.FF_DECORATIVE;
			// Latin Symbol
			case PICTORIAL:
				return FontFamily.FF_MODERN;
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

	public String getFullName()
	{
		return fullName;
	}

	public byte[] getPanose()
	{
		return panose;
	}

	@Override
		public String getTypeface()
	{
		return getFamilyName();
	}

	public int getFlags()
	{
		return flags;
	}

	@Override
		public Map<String, Supplier<?>> getGenericProperties()
	{
		final Map<String, Supplier<?>> m = new LinkedHashMap<>();
		m.put("eotSize", ()->eotSize);
		m.put("fontDataSize", ()->fontDataSize);
		m.put("version", ()->version);
		m.put("flags", getBitsAsString(this::getFlags, FLAGS_MASKS, FLAGS_NAMES));
		m.put("panose.familyType", this::getPanoseFamily);
		m.put("panose.serifType", this::getPanoseSerif);
		m.put("panose.weight", this::getPanoseWeight);
		m.put("panose.proportion", this::getPanoseProportion);
		m.put("panose.contrast", this::getPanoseContrast);
		m.put("panose.stroke", this::getPanoseStroke);
		m.put("panose.armStyle", this::getPanoseArmStyle);
		m.put("panose.letterForm", this::getPanoseLetterForm);
		m.put("panose.midLine", this::getPanoseMidLine);
		m.put("panose.xHeight", this::getPanoseXHeight);
		m.put("charset", this::getCharset);
		m.put("italic", this::isItalic);
		m.put("weight", this::getWeight);
		m.put("fsType", getBitsAsString(()->fsType, FSTYPE_MASKS, FSTYPE_NAMES));
		m.put("unicodeRange1", ()->unicodeRange1);
		m.put("unicodeRange2", ()->unicodeRange2);
		m.put("unicodeRange3", ()->unicodeRange3);
		m.put("unicodeRange4", ()->unicodeRange4);
		m.put("codePageRange1", ()->codePageRange1);
		m.put("codePageRange2", ()->codePageRange2);
		m.put("checkSumAdjustment", ()->checkSumAdjustment);
		m.put("familyName", this::getFamilyName);
		m.put("styleName", this::getStyleName);
		m.put("versionName", this::getVersionName);
		m.put("fullName", this::getFullName);
		return Collections.unmodifiableMap(m);
	}

	public PanoseFamily getPanoseFamily()
	{
		return safeEnum(PanoseFamily.values(), ()->panose[0]).get();
	}

	public PanoseSerif getPanoseSerif()
	{
		return safeEnum(PanoseSerif.values(), ()->panose[1]).get();
	}

	public PanoseWeight getPanoseWeight()
	{
		return safeEnum(PanoseWeight.values(), ()->panose[2]).get();
	}

	public PanoseProportion getPanoseProportion()
	{
		return safeEnum(PanoseProportion.values(), ()->panose[3]).get();
	}

	public PanoseContrast getPanoseContrast()
	{
		return safeEnum(PanoseContrast.values(), ()->panose[4]).get();
	}

	public PanoseStroke getPanoseStroke()
	{
		return safeEnum(PanoseStroke.values(), ()->panose[5]).get();
	}

	public PanoseArmStyle getPanoseArmStyle()
	{
		return safeEnum(PanoseArmStyle.values(), ()->panose[6]).get();
	}

	public PanoseLetterForm getPanoseLetterForm()
	{
		return safeEnum(PanoseLetterForm.values(), ()->panose[7]).get();
	}

	public PanoseMidLine getPanoseMidLine()
	{
		return safeEnum(PanoseMidLine.values(), ()->panose[8]).get();
	}

	public PanoseXHeight getPanoseXHeight()
	{
		return safeEnum(PanoseXHeight.values(), ()->panose[9]).get();
	}
}
}
