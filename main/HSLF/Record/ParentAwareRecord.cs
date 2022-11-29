using System;
using System.Collections.Generic;
using System.IO;

namespace NPOI.HSLF.Record
{
	public class ParentAwareRecord : Record
	{
		public override Record[] GetChildRecords()
		{
			throw new NotImplementedException();
		}

		public override IDictionary<string, Func<object>> GetGenericProperties()
		{
			throw new NotImplementedException();
		}

		public override long GetRecordType()
		{
			throw new NotImplementedException();
		}

		public override bool IsAnAtom()
		{
			throw new NotImplementedException();
		}

		public override void WriteOut(BinaryWriter o)
		{
			throw new NotImplementedException();
		}

		public void SetParentRecord(RecordContainer br)
		{
			throw new NotImplementedException();
		}
	}
}