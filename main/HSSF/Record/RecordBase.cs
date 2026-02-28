
namespace NPOI.HSSF.Record
{
    /**
     * Common base class of {@link Record} and {@link RecordAggregate}
     * 
     * @author Josh Micich
     */
    public abstract class RecordBase
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
        public abstract int Serialize(int offset, byte[] data);

        /**
         * gives the current serialized size of the record. Should include the sid
         * and reclength (4 bytes).
         */
        public abstract int RecordSize { get; }

        //Record CloneViaReserialise();
    }
}
