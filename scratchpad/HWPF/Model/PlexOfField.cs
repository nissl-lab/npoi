using System;
using System.Collections.Generic;
using System.Text;
using NPOI.Util;

namespace NPOI.HWPF.Model
{
    public class PlexOfField
    {

        private GenericPropertyNode propertyNode;
        private FieldDescriptor fld;

        public PlexOfField(GenericPropertyNode propertyNode)
        {
            this.propertyNode = propertyNode;
            fld = new FieldDescriptor(propertyNode.Bytes);
        }

        public int FcStart
        {
            get
            {
                return propertyNode.Start;
            }
        }

        public int FcEnd
        {
            get
            {
                return propertyNode.End;
            }
        }

        public FieldDescriptor Fld
        {
            get
            {
                return fld;
            }
        }

        public override String ToString()
        {
            return string.Format("[{0}, {1}) - FLD - 0x{2}; 0x{3}",
                    FcStart, FcEnd,
                    StringUtil.ToHexString(0xff & fld.GetBoundaryType()),
                    StringUtil.ToHexString(0xff & fld.GetFlt()));
        }
    }
}