/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HWPF.Model
{
    using System;
    using NPOI.HWPF.SPRM;
    using NPOI.HWPF.UserModel;

    /**
     */
    public class SEPX : BytePropertyNode
    {

        SectionProperties sectionProperties;
        SectionDescriptor _sed;

        public SEPX(SectionDescriptor sed, int start, int end, byte[] grpprl)
            : base(start, end, new SprmBuffer(grpprl, 0))
        {

            _sed = sed;
        }

        public byte[] GetGrpprl()
        {
            if (sectionProperties != null)
            {
                byte[] grpprl = SectionSprmCompressor
                        .CompressSectionProperty(sectionProperties);
                _buf = new SprmBuffer(grpprl, 0);
            }
            return ((SprmBuffer)_buf).ToByteArray();
        }

        public SectionDescriptor GetSectionDescriptor()
        {
            return _sed;
        }

        public SectionProperties GetSectionProperties()
        {
            if (sectionProperties == null)
            {
                sectionProperties = SectionSprmUncompressor.UncompressSEP(
                        ((SprmBuffer)_buf).ToByteArray(), 0);
            }
            return sectionProperties;
        }

        public override bool Equals(Object o)
        {
            SEPX sepx = (SEPX)o;
            if (base.Equals(o))
            {
                return sepx._sed.Equals(_sed);
            }
            return false;
        }

        public override String ToString()
        {
            return "SEPX from " + Start + " to " + End;
        }
    }
}

