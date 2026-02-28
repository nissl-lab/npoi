using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSSF.Record.Drawing
{
    /// <summary>
    /// specifies the header for an entry in a property table
    /// </summary>
    public class OfficeArtFOPTEOPID
    {
        protected ushort field_opid;
        public OfficeArtFOPTEOPID(ushort value)
        {
            this.field_opid = value;
            this.OpId = (ushort)(value & 0x3FFF);
            this.IsBlipId = (value & 0x4000) != 0;
            this.IsComplex = (value & 0x8000) != 0;
        }
        public ushort FieldOpid
        {
            get { return field_opid; }
        }
        /// <summary>
        /// specifies the identifier of the property in this entry.
        /// </summary>
        public ushort OpId
        {
            get;
            private set;
        }
        /// <summary>
        /// whether the value in the op field is a BLIP identifier. 
        /// If this value equals 0x1, the value in the op field specifies the BLIP identifier 
        /// in the OfficeArtBStoreContainer record, as defined in section 2.2.20. If fComplex equals 0x1, this bit MUST be ignored.
        /// </summary>
        public bool IsBlipId
        {
            get;
            private set;
        }
        /// <summary>
        /// specifies whether this property is a complex property. 
        /// If this value equals 0x1, the op field specifies the size of the data for this property, rather than the property data itself.
        /// </summary>
        public bool IsComplex
        {
            get;
            private set;
        }
    }
}
