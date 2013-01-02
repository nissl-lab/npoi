using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSSF.Record
{
    public interface BiffHeaderInput
    {
        /**
         * Read an unsigned short from the stream without decrypting
         */
        int ReadRecordSID();
        /**
         * Read an unsigned short from the stream without decrypting
         */
        int ReadDataSize();

        int Available();
    }
}
