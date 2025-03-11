using System;
using System.Collections.Generic;
using System.Text;
using NPOI.Util;

namespace NPOI.HSSF.Record.Drawing
{
    public class OfficeArtFOPT
    {
        private readonly OfficeArtRecordHeader _rh;
        private readonly List<OfficeArtFOPTE> _fopt;
        private readonly Dictionary<int, OfficeArtFOPTE> dictOptions = new Dictionary<int, OfficeArtFOPTE>();
        //private byte[] complexData = Array.Empty<byte>();

        public OfficeArtFOPT(RecordInputStream ris)
        {
            _rh = new OfficeArtRecordHeader(ris);
            _fopt = new List<OfficeArtFOPTE>();
            int dataRemian = _rh.Len;
            while (dataRemian > 0)
            {
                OfficeArtFOPTE opte = new OfficeArtFOPTE(ris);
                _fopt.Add(opte);
                dataRemian -= opte.DataSize;
                dictOptions.Add(opte.Opid.OpId, opte);
            }
            
        }
        public OfficeArtFOPTE GetFillOptionElement(int opid)
        {
            if (dictOptions.TryGetValue(opid, out OfficeArtFOPTE element))
                return element;
            return null;
        }
        public int DataSize
        {
            get
            {
                int size = 0;
                foreach (OfficeArtFOPTE opte in _fopt)
                    size += opte.DataSize;
                return _rh.DataSize + size;
            }
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            _rh.Serialize(out1);
            for (int i = 0; i < _fopt.Count; i++)
            {
                _fopt[i].Serialize(out1);
            }
        }
    }
    public class OfficeArtTertiaryFOPT : OfficeArtFOPT
    {
        public OfficeArtTertiaryFOPT(RecordInputStream ris)
            : base(ris)
        {
        }
    }
}
