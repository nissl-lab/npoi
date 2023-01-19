using NPOI.POIFS.Properties;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.Record
{
	public class Notes: SheetContainer
	{
		private byte[] _header;
		private static long _type = 1008L;

		// Links to our more interesting children
		private NotesAtom notesAtom;
		private PPDrawing ppDrawing;
		private ColorSchemeAtom _colorScheme;

		/**
		 * Returns the NotesAtom of this Notes
		 */
		public NotesAtom GetNotesAtom() { return notesAtom; }
		/**
		 * Returns the PPDrawing of this Notes, which has all the
		 *  interesting data in it
		 */
		public override PPDrawing GetPPDrawing() { return ppDrawing; }


		/**
		 * Set things up, and find our more interesting children
		 */
		protected Notes(byte[] source, int start, int len)
		{
			// Grab the header
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Find our children
			_children = Record.FindChildRecords(source, start + 8, len - 8);

			// Find the interesting ones in there
			foreach (Record child in _children)
			{
				if (child is NotesAtom) {
				notesAtom = (NotesAtom)child;
			}
			//if (child is PPDrawing) {
			//	ppDrawing = (PPDrawing)child;
			//}
			if (ppDrawing != null && child is ColorSchemeAtom) {
				_colorScheme = (ColorSchemeAtom)child;
			}
		}
	}


	/**
     * We are of type 1008
     */
	public override long GetRecordType() { return _type; }

	/**
     * Write the contents of the record back, so it can be written
     *  to disk
     */
	public override void WriteOut(OutputStream _out)
	{
		WriteOut(_header [0],_header [1],_type,_children, _out);
	}

	public override ColorSchemeAtom GetColorScheme()
	{
		return _colorScheme;
	}

		public override bool IsAnAtom()
		{
			return false;
		}


		public override IDictionary<string, Func<T>> GetGenericProperties<T>()
		{
			throw new NotImplementedException();
		}
	}
}
