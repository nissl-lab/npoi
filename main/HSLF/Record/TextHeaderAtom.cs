using NPOI.HSLF.Exceptions;
using NPOI.SL.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.Record
{
	public class TextHeaderAtom: RecordAtom, ParentAwareRecord
	{
		public static long _type = RecordTypes.TextHeaderAtom.typeID;
		private byte[] _header;
		private RecordContainer parentRecord;

		/** The kind of text it is */
		private int textType;
		/** position in the owning SlideListWithText */
		private int index = -1;

		public int GetTextType() { return textType; }
		public void SetTextType(int type) { textType = type; }

		public TextPlaceholder GetTextTypeEnum()
		{
			return TextPlaceholder.fromNativeId(textType);
		}

		public void SetTextTypeEnum(TextPlaceholder placeholder)
		{
			textType = placeholder.nativeId;
		}

		/**
		 * @return  0-based index of the text run in the SLWT container
		 */
		public int GetIndex() { return index; }

		/**
		 *  @param index 0-based index of the text run in the SLWT container
		 */
		public void SetIndex(int index) { this.index = index; }
		//@Override
		public RecordContainer GetParentRecord() { return parentRecord; }
		//@Override
		public void SetParentRecord(RecordContainer record) { this.parentRecord = record; }

		/* *************** record code follows ********************** */

		/**
		 * For the TextHeader Atom
		 */
		public TextHeaderAtom(byte[] source, int start, int len)
		{
			// Sanity Checking - we're always 12 bytes long
			if (len < 12)
			{
				len = 12;
				if (source.Length - start < 12)
				{
					throw new HSLFException("Not enough data to form a TextHeaderAtom (always 12 bytes long) - found " + (source.Length - start));
				}
			}

			// Get the header
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Grab the type
			textType = LittleEndian.GetInt(source, start + 8);
		}

		/**
		 * Create a new TextHeader Atom, for an unknown type of text
		 */
		public TextHeaderAtom()
		{
			_header = new byte[8];
			LittleEndian.PutUShort(_header, 0, 0);
			LittleEndian.PutUShort(_header, 2, (int)_type);
			LittleEndian.PutInt(_header, 4, 4);

			textType = (int)TextPlaceholderEnum.OTHER;
		}

		/**
		 * We are of type 3999
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

        // Write out our type
        WriteLittleEndian(textType, _out);
	}

	//@Override
	public override IDictionary<string, Func<T>> GetGenericProperties<T>()
	{
		return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
			"index", () => GetIndex(),
			"textType", GetTextTypeEnum
		);
	}
}
}
