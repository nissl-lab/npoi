using NPOI.Common.UserModel;
using NPOI.HSLF.Record;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.HSLF.Record
{
	public class TextSpecInfoRun : GenericRecord
	{
		/**
     * A enum that specifies the spelling status of a run of text.
     */
		public class SpellInfoEnum
		{
			/** the text is spelled incorrectly. */
			public static readonly SpellInfoEnum error = new SpellInfoEnum(new BitField(1));
			/** the text needs rechecking. */
			public static readonly SpellInfoEnum clean = new SpellInfoEnum(new BitField(2));
			/** the text has a grammar error. */
			public static readonly SpellInfoEnum grammar = new SpellInfoEnum(new BitField(4));
			/** the text is spelled correct */
			public static readonly SpellInfoEnum correct = new SpellInfoEnum(new BitField(0));

			public BitField bitField;

			SpellInfoEnum(BitField bitField)
			{
				this.bitField = bitField;
			}
		}

		/** A bit that specifies whether the spellInfo field exists. */
		private static BitField spellFld = BitFieldFactory.GetInstance(0X00000001);
		/** A bit that specifies whether the lid field exists. */
		private static BitField langFld = BitFieldFactory.GetInstance(0X00000002);
		/** A bit that specifies whether the altLid field exists. */
		private static BitField altLangFld = BitFieldFactory.GetInstance(0X00000004);
		// unused1, unused2 - Undefined and MUST be ignored.
		/** A bit that specifies whether the pp10runid, reserved3, and grammarError fields exist. */
		private static BitField pp10extFld = BitFieldFactory.GetInstance(0X00000020);
		/** A bit that specifies whether the bidi field exists. */
		private static BitField bidiFld = BitFieldFactory.GetInstance(0X00000040);
		// unused3 - Undefined and MUST be ignored.
		// reserved1 - MUST be zero and MUST be ignored.
		/** A bit that specifies whether the smartTags field exists. */
		private static BitField smartTagFld = BitFieldFactory.GetInstance(0X00000200);
		// reserved2 - MUST be zero and MUST be ignored.

		/**
		 * An optional unsigned integer that specifies an identifier for a character
		 * run that contains StyleTextProp11 data. It MUST exist if and only if pp10ext is TRUE.
		 **/
		private static BitField pp10runidFld = BitFieldFactory.GetInstance(0X0000000F);
		// reserved3 - An optional unsigned integer that MUST be zero, and MUST be ignored. It
		// MUST exist if and only if fPp10ext is TRUE.
		/**
		 * An optional bit that specifies a grammar error. It MUST exist if and
		 * only if fPp10ext is TRUE.
		 **/
		private static BitField grammarErrorFld = BitFieldFactory.GetInstance(unchecked((int)0X80000000));

		private static int[] FLAGS_MASKS = {
		0X00000001, 0X00000002, 0X00000004, 0X00000020, 0X00000040, 0X00000200,
	};

		private static String[] FLAGS_NAMES = {
		"SPELL", "LANG", "ALT_LANG", "PP10_EXT", "BIDI", "SMART_TAG"
	};

		//Length of special info run.
		private int length;

		//Special info mask of this run;
		private int mask;

		// info fields as indicated by the mask.
		// -1 means the bit is not set

		/**
		 * An optional SpellingFlags structure that specifies the spelling status of this
		 * text. It MUST exist if and only if spell is TRUE.
		 * The spellInfo.grammar sub-field MUST be zero.
		 * <br>
		 * error (1 bit): A bit that specifies whether the text is spelled incorrectly.<br>
		 * clean (1 bit): A bit that specifies whether the text needs rechecking.<br>
		 * grammar (1 bit): A bit that specifies whether the text has a grammar error.<br>
		 * reserved (13 bits): MUST be zero and MUST be ignored.
		 */
		private short spellInfo = -1;

		/**
		 * An optional TxLCID that specifies the language identifier of this text.
		 * It MUST exist if and only if lang is TRUE.
		 * <br>
		 * 0x0000 = No language.<br>
		 * 0x0013 = Any Dutch language is preferred over non-Dutch languages when proofing the text.<br>
		 * 0x0400 = No proofing is performed on the text.<br>
		 * &gt; 0x0400 = A valid LCID as specified by [MS-LCID].
		 */
		private short langId = -1;

		/**
		 * An optional TxLCID that specifies the alternate language identifier of this text.
		 * It MUST exist if and only if altLang is TRUE.
		 */
		private short altLangId = -1;

		/**
		 * An optional signed integer that specifies whether the text contains bidirectional
		 * characters. It MUST exist if and only if fBidi is TRUE.
		 * 0x0000 = Contains no bidirectional characters,
		 * 0x0001 = Contains bidirectional characters.
		 */
		private short bidi = -1;

		private int pp10extMask = -1;
		private byte[] smartTagsBytes;

		/**
		 * Inits a TextSpecInfoRun with default values
		 *
		 * @param len the length of the one and only run
		 */
		public TextSpecInfoRun(int len)
		{
			SetLength(len);
			SetLangId((short)0);
		}

		public TextSpecInfoRun(LittleEndianByteArrayInputStream source)
		{
			length = source.ReadInt();
			mask = source.ReadInt();
			if (spellFld.IsSet(mask))
			{
				spellInfo = source.ReadShort();
			}
			if (langFld.IsSet(mask))
			{
				langId = source.ReadShort();
			}
			if (altLangFld.IsSet(mask))
			{
				altLangId = source.ReadShort();
			}
			if (bidiFld.IsSet(mask))
			{
				bidi = source.ReadShort();
			}
			if (pp10extFld.IsSet(mask))
			{
				pp10extMask = source.ReadInt();
			}
			if (smartTagFld.IsSet(mask))
			{
				// An unsigned integer specifies the count of items in rgSmartTagIndex.
				int count = source.ReadInt();
				smartTagsBytes = IOUtils.SafelyAllocate(4 + count * 4L, RecordAtom.GetMaxRecordLength());
				LittleEndian.PutInt(smartTagsBytes, 0, count);
				// An array of SmartTagIndex that specifies the indices.
				// The count of items in the array is specified by count.
				source.ReadFully(smartTagsBytes, 4, count * 4);
			}
		}

		/**
		 * Write the contents of the record back, so it can be written
		 * to disk
		 *
		 * @param out the output stream to write to.
		 * @throws java.io.IOException if an error occurs.
		 */
		public void WriteOut(OutputStream _out)
		{
			byte[]
		   buf = new byte[4];
			LittleEndian.PutInt(buf, 0, length);
			_out.Write(buf);
			LittleEndian.PutInt(buf, 0, mask);
			_out.Write(buf);
			Object[] flds = {
				spellFld, spellInfo, "spell info",
				langFld, langId, "lang id",
				altLangFld, altLangId, "alt lang id",
				bidiFld, bidi, "bidi",
				pp10extFld, pp10extMask, "pp10 extension field",
				smartTagFld, smartTagsBytes, "smart tags"
};

			for (int i = 0; i < flds.Length - 1; i += 3)
			{
				BitField fld = (BitField)flds[i + 0];
				Object valO = flds[i + 1];
				if (!fld.IsSet(mask))
				{
					continue;
				}
				bool valid;
				if (valO is byte[])
				{
					byte[] bufB = (byte[])valO;
					valid = bufB.Length > 0;
					_out.Write(bufB);
				}
				else if (valO is int)
				{
					int valI = ((int)valO);
					valid = (valI != -1);
					LittleEndian.PutInt(buf, 0, valI);
					_out.Write(buf);
				}
				else if (valO is short)
				{
					short valS = ((short)valO);
					valid = (valS != -1);
					LittleEndian.PutShort(buf, 0, valS);
					_out.Write(buf, 0, 2);
				}
				else
				{
					valid = false;
				}
				if (!valid)
				{
					Object fval = (i + 2) < flds.Length ? flds[i + 2] : null;
					throw new IOException(fval + " is activated, but its value is invalid");
				}
			}
		}

		/**
		 * @return Spelling status of this text. null if not defined.
		 */
		public SpellInfoEnum GetSpellInfo()
		{
			if (spellInfo == -1)
			{
				return null;
			}
			foreach (SpellInfoEnum si in new SpellInfoEnum[] { SpellInfoEnum.clean, SpellInfoEnum.error, SpellInfoEnum.grammar })
			{
				if (si.bitField.IsSet(spellInfo))
				{
					return si;
				}
			}
			return SpellInfoEnum.correct;
		}

		/**
		 * @param spellInfo Spelling status of this text. null if not defined.
		 */
		public void SetSpellInfo(SpellInfoEnum spellInfo)
		{
			this.spellInfo = (spellInfo == null) ? (short)-1
				: (short)spellInfo.bitField.Set(0);
			mask = spellFld.SetBoolean(mask, spellInfo != null);
		}

		/**
		 * Windows LANGID for this text.
		 *
		 * @return Windows LANGID for this text, -1 if it's not set
		 */
		public short GetLangId()
		{
			return langId;
		}

		/**
		 * @param langId Windows LANGID for this text, -1 to unset
		 */
		public void SetLangId(short langId)
		{
			this.langId = langId;
			mask = langFld.SetBoolean(mask, langId != -1);
		}

		/**
		 * Alternate Windows LANGID of this text;
		 * must be a valid non-East Asian LANGID if the text has an East Asian language,
		 * otherwise may be an East Asian LANGID or language neutral (zero).
		 *
		 * @return  Alternate Windows LANGID of this text, -1 if it's not set
		 */
		public short GetAltLangId()
		{
			return altLangId;
		}

		public void SetAltLangId(short altLangId)
		{
			this.altLangId = altLangId;
			mask = altLangFld.SetBoolean(mask, altLangId != -1);
		}

		/**
		 * @return Length of special info run.
		 */
		public int GetLength()
		{
			return length;
		}

		/**
		 * @param length Length of special info run.
		 */
		public void SetLength(int length)
		{
			this.length = length;
		}

		/**
		 * @return the bidirectional characters flag. false = not bidi, true = is bidi, null = not set
		 */
		public bool GetBidi()
		{
			return (bidi == -1 ? false : bidi != 0);
		}

		/**
		 * @param bidi the bidirectional characters flag. false = not bidi, true = is bidi, null = not set
		 */
		public void SetBidi(bool bidi)
		{
			this.bidi = (bidi == null) ? (short)-1 : (short)(bidi ? 1 : 0);
			mask = bidiFld.SetBoolean(mask, bidi != null);
		}

		/**
		 * @return the unparsed smart tags
		 */
		public byte[] GetSmartTagsBytes()
		{
			return smartTagsBytes;
		}

		/**
		 * @param smartTagsBytes the unparsed smart tags, null to unset
		 */
		public void SetSmartTagsBytes(byte[] smartTagsBytes)
		{
			this.smartTagsBytes = (smartTagsBytes == null) ? null : (byte[])smartTagsBytes.Clone();
			mask = smartTagFld.SetBoolean(mask, smartTagsBytes != null);
		}

		/**
		 * @return an identifier for a character run that contains StyleTextProp11 data.
		 */
		public int GetPP10RunId()
		{
			return (pp10extMask == -1 || !pp10extFld.IsSet(mask)) ? -1 : pp10runidFld.GetValue(pp10extMask);

		}

		/**
		 * @param pp10RunId an identifier for a character run that contains StyleTextProp11 data, -1 to unset
		 */
		public void SetPP10RunId(int pp10RunId)
		{
			if (pp10RunId == -1)
			{
				pp10extMask = (GetGrammarError() == false) ? -1 : pp10runidFld.Clear(pp10extMask);
			}
			else
			{
				pp10extMask = pp10runidFld.SetValue(pp10extMask, pp10RunId);
			}
			// if both parameters are invalid, remove the extension mask
			mask = pp10extFld.SetBoolean(mask, pp10extMask != -1);
		}

		public bool GetGrammarError()
		{
			return (pp10extMask == -1 || !pp10extFld.IsSet(mask)) ? false : grammarErrorFld.IsSet(pp10extMask);
		}

		public void GetGrammarError(bool grammarError)
		{
			if (grammarError == false)
			{
				pp10extMask = (GetPP10RunId() == -1) ? -1 : grammarErrorFld.Clear(pp10extMask);
			}
			else
			{
				pp10extMask = grammarErrorFld.Set(pp10extMask);
			}
			// if both parameters are invalid, remove the extension mask
			mask = pp10extFld.SetBoolean(mask, pp10extMask != -1);
		}

		//@Override
		public IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
				"flags", GenericRecordUtil.GetBitsAsString(() => mask, FLAGS_MASKS, FLAGS_NAMES),
				"spellInfo", GetSpellInfo(),
				"langId", GetLangId(),
				"altLangId", GetAltLangId(),
				"bidi", GetBidi(),
				"pp10RunId", GetPP10RunId(),
				"grammarError", GetGrammarError(),
				"smartTags", GetSmartTagsBytes());

		}

		public RecordTypes GetGenericRecordType()
		{
			throw new NotImplementedException();
		}

		public IList<GenericRecord> GetGenericChildren()
		{
			throw new NotImplementedException();
		}
	}
}