using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.HSLF.Record
{
	public class UnknownRecordPlaceholder : RecordAtom
	{
		//arbitrarily selected; may need to increase
		private static int MAX_RECORD_LENGTH = 20_000_000;

		private static byte[] _contents;
		private static long _type;

		
		/**
		 * Create a new holder for a record we don't grok
		 */
		protected UnknownRecordPlaceholder(byte[] source, int start, int len)
		{
			// Sanity Checking - including whole header, so treat
			//  length as based of 0, not 8 (including header size based)
			if (len < 0) { len = 0; }

			// Treat as an atom, grab and hold everything
			_contents = IOUtils.SafelyClone(source, start, len, MAX_RECORD_LENGTH);
			_type = LittleEndian.GetUShort(_contents, 2);
		}

		public static UnknownRecordPlaceholder GetInstance(byte[] source, int start, int len)
		{
			return new UnknownRecordPlaceholder(source, start, len);
		}

		/**
		 * Return the value we were given at creation
		 */
		public override long GetRecordType()
		{
			return _type;
		}

		/**
		 * Return the value as enum we were given at creation
		 */
		public RecordTypes GetRecordTypeEnum()
		{
			return RecordTypes.ForTypeID((int)_type);
		}

		/**
		 * Write the contents of the record back, so it can be written
		 *  to disk
		 */
		public override void WriteOut(OutputStream _out)
		{
			_out.Write(_contents);
		}

		//@Override
		public override IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties("contents", () =>_contents);
		}
		
	}
}
