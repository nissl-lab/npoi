
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HWPF.SPRM
{
    using System;
    using System.Collections;
    using NPOI.HWPF.UserModel;
    using NPOI.Util;

    public class TableSprmCompressor
    {
        public TableSprmCompressor()
        {
        }
        public static byte[] CompressTableProperty(TableProperties newTAP)
        {
            int size = 0;
            ArrayList sprmList = new ArrayList();

            if (newTAP.GetJc() != 0)
            {
                size += SprmUtils.AddSprm((short)0x5400, newTAP.GetJc(), null, sprmList);
            }
            if (newTAP.GetFCantSplit())
            {
                size += SprmUtils.AddSprm((short)0x3403, 1, null, sprmList);
            }
            if (newTAP.GetFTableHeader())
            {
                size += SprmUtils.AddSprm((short)0x3404, 1, null, sprmList);
            }
            byte[] brcBuf = new byte[6 * BorderCode.SIZE];
            int offset = 0;
            newTAP.GetBrcTop().Serialize(brcBuf, offset);
            offset += BorderCode.SIZE;
            newTAP.GetBrcLeft().Serialize(brcBuf, offset);
            offset += BorderCode.SIZE;
            newTAP.GetBrcBottom().Serialize(brcBuf, offset);
            offset += BorderCode.SIZE;
            newTAP.GetBrcRight().Serialize(brcBuf, offset);
            offset += BorderCode.SIZE;
            newTAP.GetBrcHorizontal().Serialize(brcBuf, offset);
            offset += BorderCode.SIZE;
            newTAP.GetBrcVertical().Serialize(brcBuf, offset);
            byte[] compare = new byte[6 * BorderCode.SIZE];
            if (!Arrays.Equals(brcBuf, compare))
            {
                size += SprmUtils.AddSprm(unchecked((short)0xD605), 0, brcBuf, sprmList);
            }
            if (newTAP.GetDyaRowHeight() != 0)
            {
                size += SprmUtils.AddSprm(unchecked((short)0x9407), newTAP.GetDyaRowHeight(), null, sprmList);
            }
            if (newTAP.GetItcMac() > 0)
            {
                int itcMac = newTAP.GetItcMac();
                byte[] buf = new byte[1 + (LittleEndianConsts.SHORT_SIZE * (itcMac + 1)) + (TableCellDescriptor.SIZE * itcMac)];
                buf[0] = (byte)itcMac;

                short[] dxaCenters = newTAP.GetRgdxaCenter();
                for (int x = 0; x < dxaCenters.Length; x++)
                {
                    LittleEndian.PutShort(buf, 1 + (x * LittleEndianConsts.SHORT_SIZE),
                                          dxaCenters[x]);
                }

                TableCellDescriptor[] cellDescriptors = newTAP.GetRgtc();
                for (int x = 0; x < cellDescriptors.Length; x++)
                {
                    cellDescriptors[x].Serialize(buf,
                      1 + ((itcMac + 1) * LittleEndianConsts.SHORT_SIZE) + (x * TableCellDescriptor.SIZE));
                }
                size += SprmUtils.AddSpecialSprm(unchecked((short)0xD608), buf, sprmList);

                //      buf = new byte[(itcMac * ShadingDescriptor.SIZE) + 1];
                //      buf[0] = (byte)itcMac;
                //      ShadingDescriptor[] shds = newTAP.GetRgshd();
                //      for (int x = 0; x < itcMac; x++)
                //      {
                //        shds[x].Serialize(buf, 1 + (x * ShadingDescriptor.SIZE));
                //      }
                //      size += SprmUtils.AddSpecialSprm((short)0xD609, buf, sprmList);
            }
            if (newTAP.GetTlp() != 0)
            {
                size += SprmUtils.AddSprm((short)0x740a, newTAP.GetTlp(), null, sprmList);
            }

            return SprmUtils.GetGrpprl(sprmList, size);
        }
    }
}