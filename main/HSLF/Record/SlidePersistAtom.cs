using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;
using static ICSharpCode.SharpZipLib.Zip.ExtendedUnixData;

namespace NPOI.HSLF.Record
{
	public class SlidePersistAtom: RecordAtom
	{
		//arbitrarily selected; may need to increase
		private static  int MAX_RECORD_LENGTH = 32;

		private static  long _type = 1011L;
		private static  int HAS_SHAPES_OTHER_THAN_PLACEHOLDERS = 4;

		private static  int[] FLAGS_MASKS = { HAS_SHAPES_OTHER_THAN_PLACEHOLDERS };
		private static  String[] FLAGS_NAMES = { "HAS_SHAPES_OTHER_THAN_PLACEHOLDERS" };


    private  byte[] _header;

		/** Slide reference ID. Should correspond to the PersistPtr "sheet ID" of the matching slide/notes record */
		private int refID;
		private int flags;

		/** Number of placeholder texts that will follow in the SlideListWithText */
		private int numPlaceholderTexts;
		/** The internal identifier (256+), which is used to tie slides and notes together */
		private int slideIdentifier;
		/** Reserved fields. Who knows what they do */
		private byte[] reservedFields;

		public int GetRefID()
		{
			return refID;
		}

		public int GetSlideIdentifier()
		{
			return slideIdentifier;
		}

		public int GetNumPlaceholderTexts()
		{
			return numPlaceholderTexts;
		}

		public bool GetHasShapesOtherThanPlaceholders()
		{
			return (flags & HAS_SHAPES_OTHER_THAN_PLACEHOLDERS) != 0;
		}

		// Only set these if you know what you're doing!
		public void SetRefID(int id)
		{
			refID = id;
		}
		public void SetSlideIdentifier(int id)
		{
			slideIdentifier = id;
		}

		/* *************** record code follows ********************** */

		/**
		 * For the SlidePersist Atom
		 */
		protected SlidePersistAtom(byte[] source, int start, int len)
		{
			// Sanity Checking
			if (len < 8) { len = 8; }

			// Get the header
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Grab the reference ID
			refID = LittleEndian.GetInt(source, start + 8);

			// Next up is a set of flags, but only bit 3 is used!
			flags = LittleEndian.GetInt(source, start + 12);

			// Now the number of Placeholder Texts
			numPlaceholderTexts = LittleEndian.GetInt(source, start + 16);

			// Last useful one is the unique slide identifier
			slideIdentifier = LittleEndian.GetInt(source, start + 20);

			// ly you have typically 4 or 8 bytes of reserved fields,
			//  all zero running from 24 bytes in to the end
			reservedFields = IOUtils.SafelyClone(source, start + 24, len - 24, MAX_RECORD_LENGTH);
		}

		/**
		 * Create a new SlidePersistAtom, for use with a new Slide
		 */
		public SlidePersistAtom()
		{
			_header = new byte[8];
			LittleEndian.PutUShort(_header, 0, 0);
			LittleEndian.PutUShort(_header, 2, (int)_type);
			LittleEndian.PutInt(_header, 4, 20);

			flags = HAS_SHAPES_OTHER_THAN_PLACEHOLDERS;
			reservedFields = new byte[4];
		}

		/**
		 * We are of type 1011
		 */
		public override long GetRecordType() { return _type; }

		/**
		 * Write the contents of the record back, so it can be written
		 *  to disk
		 */
		public override void WriteOut(OutputStream _out)
		{
        // Header - size or type unchanged
        _out.Write(_header);

        // Write out our fields
        WriteLittleEndian(refID, _out);
		WriteLittleEndian(flags, _out);
		WriteLittleEndian(numPlaceholderTexts, _out);
		WriteLittleEndian(slideIdentifier, _out);
        _out.Write(reservedFields);
	}

	//@Override
	public override IDictionary<string, Func<T>> GetGenericProperties<T>()
	{
		return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
			"refID",()=> GetRefID(),
			"flags", GenericRecordUtil.GetBitsAsString(()=>flags, FLAGS_MASKS, FLAGS_NAMES),
			"numPlaceholderTexts", ()=> GetNumPlaceholderTexts(),
			"slideIdentifier", () => GetSlideIdentifier()
		);
	}
}
}
