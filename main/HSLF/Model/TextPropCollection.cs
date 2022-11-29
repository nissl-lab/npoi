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
using NPOI.Common.UserModel;
using NPOI.Common;
using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSLF.Exceptions;
using NPOI.HSLF.Record;
using NPOI.Util;
using System.IO;
using System.Linq;

namespace NPOI.HSLF.Model
{
	/**
 * For a given run of characters, holds the properties (which could
 *  be paragraph properties or character properties).
 * Used to hold the number of characters affected, the list of active
 *  properties, and the indent level if required.
 */
	public class TextPropCollection : GenericRecord, IDuplicatable<TextPropCollection>
	{
		/** All the different kinds of paragraph properties we might handle */
		private static TextProp[] paragraphTextPropTypes = {
        // TextProp order is according to 2.9.20 TextPFException,
        // bitmask order can be different
        new ParagraphFlagsTextProp(),
		new TextProp(2, 0x80, "bullet.char"),
		new TextProp(2, 0x10, "bullet.font"),
		new TextProp(2, 0x40, "bullet.size"),
		new TextProp(4, 0x20, "bullet.color"),
		new TextAlignmentProp(),
		new TextProp(2, 0x1000, "linespacing"),
		new TextProp(2, 0x2000, "spacebefore"),
		new TextProp(2, 0x4000, "spaceafter"),
		new TextProp(2, 0x100, "text.offset"), // left margin
        // 0x200 - Undefined and MUST be ignored
        new TextProp(2, 0x400, "bullet.offset"), // indent
        new TextProp(2, 0x8000, "defaultTabSize"),
		new HSLFTabStopPropCollection(), // tabstops size is variable!
        new FontAlignmentProp(),
		new WrapFlagsTextProp(),
		new TextProp(2, 0x200000, "textDirection"),
        // 0x400000 MUST be zero and MUST be ignored
        new TextProp(0, 0x800000, "bullet.blip"), // TODO: check size
        new TextProp(0, 0x1000000, "bullet.scheme"), // TODO: check size
        new TextProp(0, 0x2000000, "hasBulletScheme"), // TODO: check size
        // 0xFC000000 MUST be zero and MUST be ignored
    };

		/** All the different kinds of character properties we might handle */
		private static TextProp[] characterTextPropTypes = new TextProp[] {
		new TextProp(0, 0x100000, "pp10ext"),
		new TextProp(0, 0x1000000, "newAsian.font.index"), // A bit that specifies whether the newEAFontRef field of the TextCFException10 structure that contains this CFMasks exists.
        new TextProp(0, 0x2000000, "cs.font.index"), // A bit that specifies whether the csFontRef field of the TextCFException10 structure that contains this CFMasks exists.
        new TextProp(0, 0x4000000, "pp11ext"), // A bit that specifies whether the pp11ext field of the TextCFException10 structure that contains this CFMasks exists.
        new CharFlagsTextProp(),
		new TextProp(2, 0x10000, "font.index"),
		new TextProp(2, 0x200000, "asian.font.index"),
		new TextProp(2, 0x400000, "ansi.font.index"),
		new TextProp(2, 0x800000, "symbol.font.index"),
		new TextProp(2, 0x20000, "font.size"),
		new TextProp(4, 0x40000, "font.color"),
		new TextProp(2, 0x80000, "superscript")
	};

		public enum TextPropType
		{
			paragraph, character
		}

		private int charactersCovered;

		// indentLevel is only valid for paragraph collection
		// if it's set to -1, it must be omitted - see 2.9.36 TextMasterStyleLevel
		private short indentLevel;
		private Dictionary<String, TextProp> textProps = new Dictionary<string, TextProp>();
		private int maskSpecial;
		private TextPropType textPropType;

		/**
		 * Create a new collection of text properties (be they paragraph
		 *  or character) which will be groked via a subsequent call to
		 *  buildTextPropList().
		 */
		public TextPropCollection(int charactersCovered, TextPropType textPropType)
		{
			this.charactersCovered = charactersCovered;
			this.textPropType = textPropType;
		}

		public TextPropCollection(TextPropCollection other)
		{
			charactersCovered = other.charactersCovered;
			indentLevel = other.indentLevel;
			maskSpecial = other.maskSpecial;
			textPropType = other.textPropType;
			foreach (var item in other.textProps)
			{
				textProps.Add(item.Key, item.Value.Copy());
			}
		}

		public int GetSpecialMask()
		{
			return maskSpecial;
		}

		/** Fetch the number of characters this styling applies to */
		public int GetCharactersCovered()
		{
			return charactersCovered;
		}

		/** Fetch the TextProps that define this styling in the record order */
		public List<TextProp> GetTextPropList()
		{
			List<TextProp> orderedList = new List<TextProp>();
			foreach (TextProp potProp in GetPotentialProperties())
			{
				TextProp textProp = textProps[potProp.GetName()];
				if (textProp != null)
				{
					orderedList.Add(textProp);
				}
			}
			return orderedList;
		}

		/** Fetch the TextProp with this name, or null if it isn't present */
		//@SuppressWarnings("unchecked")

		public T FindByName<T>(string textPropName) where T : TextProp
		{
			return (T)textProps[textPropName];
		}

		//@SuppressWarnings("unchecked")

		public T RemoveByName<T>(string name) where T : TextProp
		{
			var result = (T)textProps[name];
			textProps.Remove(name);
			return result;
		}

		public TextPropType GetTextPropType()
		{
			return textPropType;
		}

		private TextProp[] GetPotentialProperties()
		{
			return (textPropType == TextPropType.paragraph) ? paragraphTextPropTypes : characterTextPropTypes;
		}

		/**
		 * Checks the paragraph or character properties for the given property name.
		 * Throws a HSLFException, if the name doesn't belong into this set of properties
		 *
		 * @param name the property name
		 * @return if found, the property template to copy from
		 */
		//@SuppressWarnings("unchecked")

		private T ValidatePropName<T>(string name) where T : TextProp
		{
			foreach (TextProp tp in GetPotentialProperties())
			{
				if (tp.GetName().Equals(name))
				{
					return (T)tp;
				}
			}
			String errStr =
				"No TextProp with name " + name + " is defined to add from. " +
				"Character and paragraphs have their own properties/names.";
			throw new HSLFException(errStr);
		}

		/** Add the TextProp with this name to the list */
		//@SuppressWarnings("unchecked")

		public T AddWithName<T>(string name) where T : TextProp
		{
			// Find the base TextProp to base on
			T existing = FindByName<T>(name);
			if (existing != null) return existing;

			// Add a copy of this property
			T textProp = (T)ValidatePropName<T>(name).Copy();
			textProps.Add(name, textProp);
			return textProp;
		}

		/**
		 * Add the property at the correct position. Replaces an existing property with the same name.
		 *
		 * @param textProp the property to be added
		 */
		public void AddProp(TextProp textProp)
		{
			if (textProp == null)
			{
				throw new HSLFException("TextProp must not be null");
			}

			string propName = textProp.GetName();
			ValidatePropName<TextProp>(propName);

			textProps.Add(propName, textProp);
		}

		/**
		 * For an existing set of text properties, build the list of
		 *  properties coded for in a given run of properties.
		 * @return the number of bytes that were used encoding the properties list
		 */
		public int BuildTextPropList(int containsField, byte[] data, int dataOffset)
		{
			int bytesPassed = 0;

			// For each possible entry, see if we match the mask
			// If we do, decode that, save it, and shuffle on
			foreach (TextProp tp in GetPotentialProperties())
			{
				// Check there's still data left to read

				// Check if this property is found in the mask
				if ((containsField & tp.GetMask()) != 0)
				{
					if (dataOffset + bytesPassed >= data.Length)
					{
						// Out of data, can't be any more properties to go
						// remember the mask and return
						maskSpecial |= tp.GetMask();
						return bytesPassed;
					}

					// Bingo, data contains this property
					TextProp prop = tp.Copy();
					int val = 0;
					if (prop is HSLFTabStopPropCollection)
					{
						((HSLFTabStopPropCollection)prop).ParseProperty(data, dataOffset + bytesPassed);
					}
					else if (prop.GetSize() == 2)
					{
						val = LittleEndian.GetShort(data, dataOffset + bytesPassed);
					}
					else if (prop.GetSize() == 4)
					{
						val = LittleEndian.GetInt(data, dataOffset + bytesPassed);
					}
					else if (prop.GetSize() == 0)
					{
						//remember "special" bits.
						maskSpecial |= tp.GetMask();
						continue;
					}

					if (prop is BitMaskTextProp)
					{
						((BitMaskTextProp)prop).SetValueWithMask(val, containsField);
					}
					else if (!(prop is HSLFTabStopPropCollection))
					{
						prop.SetValue(val);
					}
					bytesPassed += prop.GetSize();
					AddProp(prop);
				}
			}

			// Return how many bytes were used
			return bytesPassed;
		}

		/**
		 * Clones the given text properties
		 */
		//@Override
		public TextPropCollection Copy()
		{
			return new TextPropCollection(this);
		}

		/**
		 * Update the size of the text that this set of properties
		 *  applies to
		 */
		public void UpdateTextSize(int textSize)
		{
			charactersCovered = textSize;
		}

		/**
		* Writes out to disk the header, and then all the properties
*/
		public void WriteOut(OutputStream o)
		{
			WriteOut(o, false);
		}

		/**
		 * Writes out to disk the header, and then all the properties
		 */
		public void WriteOut(OutputStream o, bool isMasterStyle)
		{
			if (!isMasterStyle)
			{
				// First goes the number of characters we affect
				// MasterStyles don't have this field
				Record.Record.WriteLittleEndian(charactersCovered, o);
			}

			// Then we have the indentLevel field if it's a paragraph collection
			if (textPropType == TextPropType.paragraph && indentLevel > -1)
			{
				Record.Record.WriteLittleEndian(indentLevel, o);
			}

			// Then the mask field
			int mask = maskSpecial;
			foreach (TextProp textProp in textProps.Values)
			{
				mask |= textProp.GetWriteMask();
			}
			Record.Record.WriteLittleEndian(mask, o);

			// Then the contents of all the properties
			foreach (TextProp textProp in GetTextPropList())
			{
				int val = textProp.GetValue();
				if (textProp is BitMaskTextProp && textProp.GetWriteMask() == 0)
				{
					// don't add empty properties, as they can't be recognized while reading
					continue;
				}
				else if (textProp.GetSize() == 2)
				{
					Record.Record.WriteLittleEndian((short)val, o);
				}
				else if (textProp.GetSize() == 4)
				{
					Record.Record.WriteLittleEndian(val, o);
				}
				else if (textProp is HSLFTabStopPropCollection)
				{
					((HSLFTabStopPropCollection)textProp).WriteProperty(o);
				}
			}
		}

		public short GetIndentLevel()
		{
			return indentLevel;
		}

		public void setIndentLevel(short indentLevel)
		{
			if (textPropType == TextPropType.character)
			{
				throw new InvalidOperationException("trying to set an indent on a character collection.");
			}
			this.indentLevel = indentLevel;
		}

		public override int GetHashCode()
		{
			return Tuple.Create(charactersCovered, maskSpecial, indentLevel, textProps).GetHashCode();
		}
		/**
		 * compares most properties apart of the covered characters length
		 */
		public override bool Equals(Object other)
		{
			if (this == other) return true;
			if (other == null) return false;
			if (this.GetType() != other.GetType()) return false;

			TextPropCollection o = (TextPropCollection)other;
			if (o.maskSpecial != this.maskSpecial || o.indentLevel != this.indentLevel)
			{
				return false;
			}

			return textProps.Equals(o.textProps);
		}

		public override string ToString()
		{
			StringBuilder _out = new StringBuilder();
			_out.Append("  chars covered: ").Append(GetCharactersCovered());
			_out.Append("  special mask flags: 0x").Append(HexDump.ToHex(GetSpecialMask())).Append("\n");
			if (textPropType == TextPropType.paragraph)
			{
				_out.Append("  indent level: ").Append(GetIndentLevel()).Append("\n");
			}
			foreach (TextProp p in GetTextPropList())
			{
				_out.Append("    ");
				_out.Append(p.ToString());
				_out.Append("\n");
				if (p is BitMaskTextProp)
				{
					BitMaskTextProp bm = (BitMaskTextProp)p;
					int i = 0;
					foreach (string s in bm.GetSubPropNames())
					{
						if (bm.GetSubPropMatches()[i])
						{
							_out.Append("          ").Append(s).Append(" = ").Append(bm.GetSubValue(i)).Append("\n");
						}
						i++;
					}
				}
			}

			_out.Append("  bytes that would be written: \n");

			try
			{
				MemoryStream baos = new MemoryStream();
				WriteOut((OutputStream)baos);
				byte[] b = baos.ToArray();
				_out.Append(HexDump.Dump(b, 0, 0));
			}
			catch (IOException e)
			{
				//LOG.atError().withThrowable(e).log("can't dump TextPropCollection");
			}

			return _out.ToString();
		}



		//@Override
		public IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			IDictionary<string, Func<object>> m = new Dictionary<string, Func<object>>();
			m.Add("charactersCovered", () => GetCharactersCovered());
			m.Add("indentLevel", () => GetIndentLevel());
			foreach (var item in textProps)
			{
				m.Add(item.Key, () => item.Value);
			}
			m.Add("maskSpecial", () => GetSpecialMask());
			m.Add("textPropType", () => GetTextPropType());
			return (IDictionary<string, Func<T>>)m;
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
