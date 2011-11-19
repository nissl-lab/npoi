using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Model
{
    public class FSPADocumentPart
    {
        public static int HEADER = FIBFieldHandler.PLCSPAHDR;

        public static int MAIN = FIBFieldHandler.PLCSPAMOM;

        private int fibFieldsField;

        private FSPADocumentPart(int fibHandlerField)
        {
            this.fibFieldsField = fibHandlerField;
        }

        public int FibFieldsField
        {
            get { return fibFieldsField; }
        }
    }
}
