using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.HSLF.Record
{
	/**
 * Abstract class which all container records will extend. Providers
 *  helpful methods for writing child records out to disk
 */

	public abstract class RecordContainer : Record
	{

		protected Record[] _children;

		/**
		 * Return any children
		 */
		//@Override
		public override Record[] GetChildRecords() { return _children; }

		/**
		 * We're not an atom
		 */
		//@Override
		public bool isAnAtom() { return false; }


		/* ===============================================================
		 *                   Internal Move Helpers
		 * ===============================================================
		 */

		/**
		 * Finds the location of the given child record
		 */
		private int FindChildLocation(Record child)
		{
			int i = 0;
			foreach (Record r in _children)
			{
				if (r.Equals(child))
				{
					return i;
				}
				i++;
			}
			return -1;
		}

		/**
		 * Adds a child record, at the very end.
		 * @param newChild The child record to add
		 * @return the position of the added child
		 */
		private int AppendChild(Record newChild)
		{
			// Copy over, and pop the child in at the end
			Record[] nc = Arrays.CopyOf<Record>(_children, _children.Length + 1);
			// Switch the arrays
			nc[_children.Length] = newChild;
			_children = nc;
			return _children.Length;
		}

		/**
		 * Adds the given new Child Record at the given location,
		 *  shuffling everything from there on down by one
		 *
		 * @param newChild The record to be added as child-record.
		 * @param position The index where the child should be added, 0-based
		 */
		private void AddChildAt(Record newChild, int position)
		{
			// Firstly, have the child added in at the end
			AppendChild(newChild);

			// Now, have them moved to the right place
			MoveChildRecords((_children.Length - 1), position, 1);
		}

		/**
		 * Moves {@code number} child records from {@code oldLoc} to {@code newLoc}.
		 * @param oldLoc the current location of the records to move
		 * @param newLoc the new location for the records
		 * @param number the number of records to move
		 */
		private void MoveChildRecords(int oldLoc, int newLoc, int number)
		{
			if (oldLoc == newLoc) { return; }
			if (number == 0) { return; }

			// Check that we're not asked to move too many
			if (oldLoc + number > _children.Length)
			{
				throw new InvalidOperationException("Asked to move more records than there are!");
			}

			// Do the move
			Arrays.ArrayMoveWithin(_children, oldLoc, newLoc, number);
		}


		/**
		 * Finds the first child record of the given type,
		 *  or null if none of the child records are of the
		 *  given type. Does not descend.
		 */
		public Record FindFirstOfType(long type)
		{
			foreach (Record r in _children)
			{
				if (r.GetRecordType() == type)
				{
					return r;
				}
			}
			return null;
		}

		/**
		 * Remove a child record from this record container
		 *
		 * @param ch the child to remove
		 * @return the removed record
		 */
		public Record RemoveChild(Record ch)
		{
			Record rm = null;
			List<Record> lst = new List<Record>();
			foreach (Record r in _children)
			{
				if (r != ch)
				{
					lst.Add(r);
				}
				else
				{
					rm = r;
				}
			}
			_children = lst.ToArray();
			return rm;
		}

		/* ===============================================================
		 *                   External Move Methods
		 * ===============================================================
		 */

		/**
		 * Add a new child record onto a record's list of children.
		 *
		 * @param newChild the child record to be added
		 * @return the position of the added child within the list, i.e. the last index
		 */
		public int AppendChildRecord(Record newChild)
		{
			return AppendChild(newChild);
		}

		/**
		 * Adds the given Child Record after the supplied record
		 * @param newChild The record to add as new child.
		 * @param after The record after which the given record should be added.
		 * @return the position of the added child within the list
		 */
		public int AddChildAfter(Record newChild, Record after)
		{
			// Decide where we're going to put it
			int loc = FindChildLocation(after);
			if (loc == -1)
			{
				throw new InvalidOperationException("Asked to add a new child after another record, but that record wasn't one of our children!");
			}

			// Add one place after the supplied record
			AddChildAt(newChild, loc + 1);
			return loc + 1;
		}

		/**
		 * Adds the given Child Record before the supplied record
		 * @param newChild The record to add as new child.
		 * @param before The record before which the given record should be added.
		 * @return the position of the added child within the list
		 */
		public int AddChildBefore(Record newChild, Record before)
		{
			// Decide where we're going to put it
			int loc = FindChildLocation(before);
			if (loc == -1)
			{
				throw new InvalidOperationException("Asked to add a new child before another record, but that record wasn't one of our children!");
			}

			// Add at the place of the supplied record
			AddChildAt(newChild, loc);
			return loc;
		}

		/**
		 * Set child records.
		 *
		 * @param records   the new child records
		 */
		public void SetChildRecord(Record[] records)
		{
			this._children = (Record[])records.Clone();
		}

		/* ===============================================================
		 *                 External Serialisation Methods
		 * ===============================================================
		 */

		/**
		 * Write out our header, and our children.
		 * @param headerA the first byte of the header
		 * @param headerB the second byte of the header
		 * @param type the record type
		 * @param children our child records
		 * @param out the stream to write to
		 */
		public void WriteOut(byte headerA, byte headerB, long type, Record[] children, OutputStream _out)
		{
			// Create a UnsynchronizedByteArrayOutputStream to hold everything in
			using (MemoryStream stream = new MemoryStream())
			{
				using (BinaryWriter writer = new BinaryWriter(stream))
				{
					// Write out our header, less the size
					writer.Write(new byte[] { headerA, headerB });
					byte[] typeB = new byte[2];
					LittleEndian.PutShort(typeB, 0, (short)type);
					writer.Write(typeB);
					writer.Write(new byte[] { 0, 0, 0, 0 });

					// Write out our children
					foreach (Record aChildren in children)
					{
						aChildren.WriteOut(writer);
					}

					// Grab the bytes back
					byte[] toWrite = stream.ToArray();

					// Update our header with the size
					// Don't forget to knock 8 more off, since we don't include the header in the size
					LittleEndian.PutInt(toWrite, 4, (toWrite.Length - 8));

					// Write out the bytes
					_out.Write(toWrite);
				}
			}
		}

		/**
		 * Find the records that are parent-aware, and tell them who their parent is
		 */
		public static void HandleParentAwareRecords(RecordContainer br)
		{
			// Loop over child records, looking for interesting ones
			foreach (Record record in br.GetChildRecords())
			{
				// Tell parent aware records of their parent
				if (record is ParentAwareRecord)
				{
					((ParentAwareRecord)record).SetParentRecord(br);
				}
				// Walk on down for the case of container records
				if (record is RecordContainer)
				{
					HandleParentAwareRecords((RecordContainer)record);
				}
			}
		}

		//@Override
		public override IDictionary<string, Func<object>> GetGenericProperties()
		{
			return null;
		}
	}
}
