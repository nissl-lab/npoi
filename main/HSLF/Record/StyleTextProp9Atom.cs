using NPOI.HSLF.Model;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.Record
{
	public class StyleTextProp9Atom: RecordAtom
	{
		//arbitrarily selected; may need to increase
		private static int DEFAULT_MAX_RECORD_LENGTH = 100_000;
		private static int MAX_RECORD_LENGTH = DEFAULT_MAX_RECORD_LENGTH;

		private TextPFException9[] autoNumberSchemes;
    /** Record header. */
    private byte[] header;
		/** Record data. */
		private byte[] data;
		private short version;
		private short recordId;
		private int length;

		/**
		 * @param length the max record length allowed for StyleTextProp9Atom
		 */
		public static void SetMaxRecordLength(int length)
		{
			MAX_RECORD_LENGTH = length;
		}

		/**
		 * @return the max record length allowed for StyleTextProp9Atom
		 */
		public static int GetMaxRecordLength()
		{
			return MAX_RECORD_LENGTH;
		}

		/**
		 * Constructs the link related atom record from its
		 *  source data.
		 *
		 * @param source the source data as a byte array.
		 * @param start the start offset into the byte array.
		 * @param len the length of the slice in the byte array.
		 */
		protected StyleTextProp9Atom(byte[] source, int start, int len)
		{
			// Get the header.
			List<TextPFException9> schemes = new List<TextPFException9>();
			header = Arrays.CopyOfRange(source, start, start + 8);
			this.version = LittleEndian.GetShort(header, 0);
			this.recordId = LittleEndian.GetShort(header, 2);
			this.length = LittleEndian.GetInt(header, 4);

			// Get the record data.
			data = IOUtils.SafelyClone(source, start + 8, len - 8, MAX_RECORD_LENGTH);
			for (int i = 0; i < data.Length;)
			{
				TextPFException9 item = new TextPFException9(data, i);
				schemes.Add(item);
				i += item.GetRecordLength();

				if (i + 4 >= data.Length)
				{
					break;
				}
				int textCfException9 = LittleEndian.GetInt(data, i);
				i += 4;
				//TODO analyze textCfException when have some test data

				if (i + 4 >= data.Length)
				{
					break;
				}
				int textSiException = LittleEndian.GetInt(data, i);
				i += 4;//TextCFException9 + SIException

				if (0 != (textSiException & 0x40))
				{
					i += 2; //skip fBidi
				}
				if (i + 4 >= data.Length)
				{
					break;
				}
			}
			this.autoNumberSchemes = schemes.ToArray();
		}

		/**
		 * Gets the record type.
		 * @return the record type.
		 */
		public override long GetRecordType() { return this.recordId; }

		public short GetVersion()
		{
			return version;
		}

		public int GetLength()
		{
			return length;
		}
		public TextPFException9[] GetAutoNumberTypes()
		{
			return this.autoNumberSchemes;
		}

		/**
		 * Write the contents of the record back, so it can be written
		 * to disk
		 *
		 * @param out the output stream to write to.
		 * @throws java.io.IOException if an error occurs.
		 */
		public override void WriteOut(OutputStream _out)
		{
        _out.Write(header);
        _out.Write(data);
	}

	/**
     * Update the text length
     *
     * @param size the text length
     */
	public void SetTextSize(int size)
	{
		LittleEndian.PutInt(data, 0, size);
	}

	/**
     * Reset the content to one info run with the default values
     * @param size  the site of parent text
     */
	public void Reset(int size)
	{
		data = new byte[10];
		// 01 00 00 00
		LittleEndian.PutInt(data, 0, size);
		// 01 00 00 00
		LittleEndian.PutInt(data, 4, 1); //mask
										 // 00 00
		LittleEndian.PutShort(data, 8, (short)0); //langId

		// Update the size (header bytes 5-8)
		LittleEndian.PutInt(header, 4, data.Length);
	}

	//@Override
	public override IDictionary<string, Func<T>> GetGenericProperties<T>()
	{
		return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
			"autoNumberSchemes", GetAutoNumberTypes
		);
	}
}
}
