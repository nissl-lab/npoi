using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HWPF.Model
{
    public enum SubdocumentType:int
    {
        MAIN= FIBLongHandler.CCPTEXT ,

        FOOTNOTE= FIBLongHandler.CCPFTN ,

        HEADER= FIBLongHandler.CCPHDD ,

        MACRO= FIBLongHandler.CCPMCR ,

        ANNOTATION= FIBLongHandler.CCPATN ,

        ENDNOTE= FIBLongHandler.CCPEDN ,

        TEXTBOX= FIBLongHandler.CCPTXBX ,

        HEADER_TEXTBOX= FIBLongHandler.CCPHDRTXBX
    }


}
