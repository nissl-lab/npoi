using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.Record
{
	public class TextCharsAtom : RecordAtom
	{
		public static long _type = RecordTypes.TextCharsAtom.typeID;

		private byte[] _header;

		/** The bytes that make up the text */
		private byte[] _text;

		/** Grabs the text. */
		public String GetText()
		{
			return StringUtil.GetFromUnicodeLE(_text);
		}

		/** Updates the text in the Atom. */
		public void SetText(String text)
		{
			// Convert to little endian unicode
			_text = IOUtils.SafelyAllocate(text.Length * 2L, GetMaxRecordLength());
			StringUtil.PutUnicodeLE(text, _text, 0);

			// Update the size (header bytes 5-8)
			LittleEndian.PutInt(_header, 4, _text.Length);
		}

		/* *************** record code follows ********************** */

		/**
		 * For the TextChars Atom
		 */
		protected TextCharsAtom(byte[] source, int start, int len)
		{
			// Sanity Checking
			if (len < 8) { len = 8; }

			// Get the header
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Grab the text
			_text = IOUtils.SafelyClone(source, start + 8, len - 8, GetMaxRecordLength());
		}
		/**
		 * Create an empty TextCharsAtom
		 */
		public TextCharsAtom()
		{
			// 0 length header
			_header = new byte[] { 0, 0, 0xA0, 0x0f, 0, 0, 0, 0 };
			// Empty text
			_text = new byte[0];
		}

		/**
		 * We are of type 4000
		 */
		//@Override
		public override long GetRecordType() { return _type; }

		/**
		 * Write the contents of the record back, so it can be written
		 *  to disk
		 */
		//@Override
		public override void WriteOut(OutputStream _out)
		{
			// Header - size or type unchanged
			_out.Write(_header);

			// Write out our text
			_out.Write(_text);
		}

		/**
		 * dump debug info; use getText() to return a string
		 * representation of the atom
		 */
		//@Override
		public override String ToString()
		{
			return GenericRecordJsonWriter.marshal(this);
		}

		//@Override
		public override IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
				"text", GetText
			);
		}
	}
}
