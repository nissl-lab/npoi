using NPOI.SL.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;
using static ICSharpCode.SharpZipLib.Zip.ExtendedUnixData;

namespace NPOI.HSLF.Record
{
	public class NotesAtom: RecordAtom
	{
		private byte[] _header;
		private static long _type = 1009L;

		private int slideID;
		private bool followMasterObjects;
		private bool followMasterScheme;
		private bool followMasterBackground;
		private byte[] reserved;


		public int GetSlideID() { return slideID; }
		public void SetSlideID(int id) { slideID = id; }

		public bool GetFollowMasterObjects() { return followMasterObjects; }
		public bool GetFollowMasterScheme() { return followMasterScheme; }
		public bool GetFollowMasterBackground() { return followMasterBackground; }
		public void SetFollowMasterObjects(bool flag) { followMasterObjects = flag; }
		public void SetFollowMasterScheme(bool flag) { followMasterScheme = flag; }
		public void SetFollowMasterBackground(bool flag) { followMasterBackground = flag; }


		/* *************** record code follows ********************** */

		/**
		 * For the Notes Atom
		 */
		protected NotesAtom(byte[] source, int start, int len)
		{
			// Sanity Checking
			if (len < 8) { len = 8; }

			// Get the header
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Get the slide ID
			slideID = LittleEndian.GetInt(source, start + 8);

			// Grok the flags, stored as bits
			int flags = LittleEndian.GetUShort(source, start + 12);
			followMasterBackground = (flags & 4) == 4;
			followMasterScheme = (flags & 2) == 2;
			followMasterObjects = (flags & 1) == 1;

			// There might be 2 more bytes, which are a reserved field
			reserved = IOUtils.SafelyClone(source, start + 14, len - 14, GetMaxRecordLength());
		}

		/**
		 * We are of type 1009
		 */
		public override long GetRecordType() { return _type; }

		/**
		 * Write the contents of the record back, so it can be written
		 *  to disk
		 */
		public override void WriteOut(OutputStream _out)
		{
        // Header
        _out.Write(_header);

        // Slide ID
        WriteLittleEndian(slideID, _out);

		// Flags
		short flags = 0;
        if(followMasterObjects)    { flags += (short) 1; }
        if(followMasterScheme)     { flags += (short) 2; }
if (followMasterBackground) { flags += (short)4; }
WriteLittleEndian(flags, _out);

        // Reserved fields
        _out.Write(reserved);
    }

    //@Override
	public override IDictionary<string, Func<T>> GetGenericProperties<T>()
{
	return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
		"slideId", ()=>GetSlideID(),
		"followMasterObjects", ()=>GetFollowMasterObjects(),
		"followMasterScheme", ()=>GetFollowMasterScheme(),
		"followMasterBackground", ()=>GetFollowMasterBackground()
	);
}
	}
}
