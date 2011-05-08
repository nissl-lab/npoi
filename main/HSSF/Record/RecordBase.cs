using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSSF.Record
{
    /**
     * Common base class of {@link Record} and {@link RecordAggregate}
     * 
     * @author Josh Micich
     */
    public interface RecordBase
    {
        /**
         * called by the class that is responsible for writing this sucker.
         * Subclasses should implement this so that their data is passed back in a
         * byte array.
         * 
         * @param offset to begin writing at
         * @param data byte array containing instance data
         * @return number of bytes written
         */
        int Serialize(int offset, byte[] data);

        /**
         * gives the current serialized size of the record. Should include the sid
         * and reclength (4 bytes).
         */
        int RecordSize{get;}

        Record CloneViaReserialise();
    }
}
