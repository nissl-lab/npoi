/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
namespace NPOI.HWPF.Model
{
    using System;
    using System.Text;
    using System.Collections;
    using NPOI.DDF;
    using System.Collections.Generic;

    /**
     * Based on AbstractEscherRecordHolder fomr HSSF.
     * 
     * @author Squeeself
     */
    public class EscherRecordHolder
    {
        protected List<EscherRecord> escherRecords = new List<EscherRecord>();

        public EscherRecordHolder()
        {

        }

        public EscherRecordHolder(byte[] data, int offset, int size)
        {
            FillEscherRecords(data, offset, size);
        }

        private void FillEscherRecords(byte[] data, int offset, int size)
        {
            IEscherRecordFactory recordFactory = new DefaultEscherRecordFactory();
            int pos = offset;
            while (pos < offset + size)
            {
                EscherRecord r = recordFactory.CreateRecord(data, pos);
                escherRecords.Add(r);
                int bytesRead = r.FillFields(data, pos, recordFactory);
                pos += bytesRead + 1; // There Is an empty byte between each top-level record in a Word doc
            }
        }

        public List<EscherRecord> EscherRecords
        {
            get
            {
                return escherRecords;
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            String nl = System.Environment.NewLine;
            if (escherRecords.Count == 0)
                buffer.Append("No Escher Records Decoded" + nl);
            for (IEnumerator iterator = escherRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)iterator.Current;
                buffer.Append(r.ToString());
            }

            return buffer.ToString();
        }

        /**
         * If we have a EscherContainerRecord as one of our
         *  children (and most top level escher holders do),
         *  then return that.
         */
        public EscherContainerRecord EscherContainer
        {
            get
            {
                for (IEnumerator it = escherRecords.GetEnumerator(); it.MoveNext(); )
                {
                    Object er = it.Current;
                    if (er is EscherContainerRecord)
                    {
                        return (EscherContainerRecord)er;
                    }
                }
                return null;
            }
        }

        /**
         * Descends into all our children, returning the
         *  first EscherRecord with the given id, or null
         *  if none found
         */
        public EscherRecord FindFirstWithId(short id)
        {
            return FindFirstWithId(id, this.EscherRecords);
        }
        private EscherRecord FindFirstWithId(short id, IList records)
        {
            // Check at our level
            for (IEnumerator it = records.GetEnumerator(); it.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)it.Current;
                if (r.RecordId == id)
                {
                    return r;
                }
            }

            // Then check our children in turn
            for (IEnumerator it = records.GetEnumerator(); it.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)it.Current;
                if (r.IsContainerRecord)
                {
                    EscherRecord found =
                        FindFirstWithId(id, r.ChildRecords);
                    if (found != null)
                    {
                        return found;
                    }
                }
            }

            // Not found in this lot
            return null;
        }

        public List<EscherRecord> GetEscherRecords()
        {
            return escherRecords;
        }


        public List<EscherContainerRecord> GetDgContainers()
        {
            List<EscherContainerRecord> dgContainers = new List<EscherContainerRecord>(
                    1);
            foreach (EscherRecord escherRecord in GetEscherRecords())
            {
                if (escherRecord.RecordId == unchecked((short)0xF002))
                {
                    dgContainers.Add((EscherContainerRecord)escherRecord);
                }
            }
            return dgContainers;
        }

        public List<EscherContainerRecord> GetDggContainers()
        {
            List<EscherContainerRecord> dggContainers = new List<EscherContainerRecord>(
                    1);
            foreach (EscherRecord escherRecord in GetEscherRecords())
            {
                if (escherRecord.RecordId == unchecked((short)0xF000))
                {
                    dggContainers.Add((EscherContainerRecord)escherRecord);
                }
            }
            return dggContainers;
        }

        public List<EscherContainerRecord> GetBStoreContainers()
        {
            List<EscherContainerRecord> bStoreContainers = new List<EscherContainerRecord>(
                    1);
            foreach (EscherContainerRecord dggContainer in GetDggContainers())
            {
                foreach (EscherRecord escherRecord in dggContainer.ChildRecords)
                {
                    if (escherRecord.RecordId == unchecked((short)0xF001))
                    {
                        bStoreContainers.Add((EscherContainerRecord)escherRecord);
                    }
                }
            }
            return bStoreContainers;
        }

        public List<EscherContainerRecord> GetSpgrContainers()
        {
            List<EscherContainerRecord> spgrContainers = new List<EscherContainerRecord>(
                    1);
            foreach (EscherContainerRecord dgContainer in GetDgContainers())
            {
                foreach (EscherRecord escherRecord in dgContainer.ChildRecords)
                {
                    if (escherRecord.RecordId == unchecked((short)0xF003))
                    {
                        spgrContainers.Add((EscherContainerRecord)escherRecord);
                    }
                }
            }
            return spgrContainers;
        }

        public List<EscherContainerRecord> GetSpContainers()
        {
            List<EscherContainerRecord> spContainers = new List<EscherContainerRecord>(
                    1);
            foreach (EscherContainerRecord spgrContainer in GetSpgrContainers())
            {
                foreach (EscherRecord escherRecord in spgrContainer.ChildRecords)
                {
                    if (escherRecord.RecordId == unchecked((short)0xF004))
                    {
                        spContainers.Add((EscherContainerRecord)escherRecord);
                    }
                }
            }
            return spContainers;
        }
    }
}