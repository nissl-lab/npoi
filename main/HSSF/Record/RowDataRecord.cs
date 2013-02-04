using System;

namespace NPOI.HSSF.Record
{
    /// <summary>
    /// this record only used for record that has name and not implemented.
    /// </summary>
    public abstract class RowDataRecord : StandardRecord
    {
        private byte[] _rawData = null;
        public RowDataRecord(RecordInputStream in1)
        {
            _rawData = in1.ReadRemainder();
        }
        protected override int DataSize
        {
            get
            {
                return _rawData.Length;
            }
        }

        public override void Serialize(NPOI.Util.ILittleEndianOutput out1)
        {
            out1.Write(_rawData);
        }

        public override short Sid
        {
            get { throw new NotImplementedException("must be implemented in sub class"); }
        }
    }
}
