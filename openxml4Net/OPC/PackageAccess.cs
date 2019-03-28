using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXml4Net.OPC
{
    /**
     * Specifies package access.
     *
     * @author Julien Chable
     * @version 1.0
     */
    public enum PackageAccess
    {
        /** Read only. Write not authorized. */
        READ,
        /** Write only. Read not authorized. */
        WRITE,
        /** Read and Write mode. */
        READ_WRITE
    }

}
