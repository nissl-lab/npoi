using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.Record
{
	public class CString : RecordAtom
	{
		private byte[] _header;

		/** The bytes that make up the text */
		private byte[] _text;

		/** Grabs the text. Never <code>null</code> */
		public String GetText()
		{
			return StringUtil.GetFromUnicodeLE(_text);
		}

		/** Updates the text in the Atom. */
		public void SetText(String text)
		{
			// Convert to little endian unicode
			_text = new byte[text.Length * 2];
			StringUtil.PutUnicodeLE(text, _text, 0);

			// Update the size (header bytes 5-8)
			LittleEndian.PutInt(_header, 4, _text.Length);
		}

		/**
		 * Grabs the count, from the first two bytes of the header.
		 * The meaning of the count is specific to the type of the parent record
		 */
		public int GetOptions()
		{
			return LittleEndian.GetShort(_header);
		}

		/**
		 * Sets the count
		 * The meaning of the count is specific to the type of the parent record
		 */
		public void SetOptions(int count)
		{
			LittleEndian.PutShort(_header, 0, (short)count);
		}

		/* *************** record code follows ********************** */

		/**
		 * For the CStrubg Atom
		 */
		protected CString(byte[] source, int start, int len)
		{
			// Sanity Checking
			if (len < 8) { len = 8; }

			// Get the header
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Grab the text
			_text = IOUtils.SafelyClone(source, start + 8, len - 8, GetMaxRecordLength());
		}
		/**
		 * Create an empty CString
		 */
		public CString()
		{
			// 0 length header
			_header = new byte[] { 0, 0, 0xBA, 0x0f, 0, 0, 0, 0 };
			// Empty text
			_text = new byte[0];
		}

		/**
		 * We are of type 4026
		 */
		public long GetRecordType()
		{
			return RecordTypes.CString.typeID;
		}

		/**
		 * Write the contents of the record back, so it can be written
		 *  to disk
		 */
		public void WriteOut(OutputStream _out)
		{
			// Header - size or type unchanged
			_out.Write(_header);

			// Write out our text
			_out.Write(_text);
		}

		/**
		 * Gets a string representation of this object, primarily for debugging.
		 * @return a string representation of this object.
		 */
		public override String ToString()
		{
			return GetText();
		}

		//@Override
		public Dictionary<string, Func<T>> GetGenericProperties()
		{
			return GenericRecordUtil.getGenericProperties("text", this.GetText);
		}
	}
}
