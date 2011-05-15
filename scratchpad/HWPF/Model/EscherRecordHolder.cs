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

    /**
     * Based on AbstractEscherRecordHolder fomr HSSF.
     * 
     * @author Squeeself
     */
    public class EscherRecordHolder
    {
        protected ArrayList escherRecords = new ArrayList();

        public EscherRecordHolder()
        {

        }

        public EscherRecordHolder(byte[] data, int offset, int size)
        {
            FillEscherRecords(data, offset, size);
        }

        private void FillEscherRecords(byte[] data, int offset, int size)
        {
            EscherRecordFactory recordFactory = new DefaultEscherRecordFactory();
            int pos = offset;
            while (pos < offset + size)
            {
                EscherRecord r = recordFactory.CreateRecord(data, pos);
                escherRecords.Add(r);
                int bytesRead = r.FillFields(data, pos, recordFactory);
                pos += bytesRead + 1; // There Is an empty byte between each top-level record in a Word doc
            }
        }

        public IList EscherRecords
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
    }
}