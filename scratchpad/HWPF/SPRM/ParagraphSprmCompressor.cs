
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

    public class ParagraphSprmCompressor
    {
        public ParagraphSprmCompressor()
        {
        }

        public static byte[] compressParagraphProperty(ParagraphProperties newPAP,
                                                       ParagraphProperties oldPAP)
        {
            ArrayList sprmList = new ArrayList();
            int size = 0;

            if (newPAP.GetJc() != oldPAP.GetJc())
            {
                size += SprmUtils.AddSprm((short)0x2403, newPAP.GetJc(), null, sprmList);
            }
            if (newPAP.GetFKeep() != oldPAP.GetFKeep())
            {
                size += SprmUtils.AddSprm((short)0x2405, newPAP.GetFKeep(), null, sprmList);
            }
            if (newPAP.GetFKeepFollow() != oldPAP.GetFKeepFollow())
            {
                size += SprmUtils.AddSprm((short)0x2406, newPAP.GetFKeepFollow(), null, sprmList);
            }
            if (newPAP.GetFPageBreakBefore() != oldPAP.GetFPageBreakBefore())
            {
                size += SprmUtils.AddSprm((short)0x2407, newPAP.GetFPageBreakBefore(), null, sprmList);
            }
            if (newPAP.GetBrcl() != oldPAP.GetBrcl())
            {
                size += SprmUtils.AddSprm((short)0x2408, newPAP.GetBrcl(), null, sprmList);
            }
            if (newPAP.GetBrcp() != oldPAP.GetBrcp())
            {
                size += SprmUtils.AddSprm((short)0x2409, newPAP.GetBrcp(), null, sprmList);
            }
            if (newPAP.GetIlvl() != oldPAP.GetIlvl())
            {
                size += SprmUtils.AddSprm((short)0x260A, newPAP.GetIlvl(), null, sprmList);
            }
            if (newPAP.GetIlfo() != oldPAP.GetIlfo())
            {
                size += SprmUtils.AddSprm((short)0x460b, newPAP.GetIlfo(), null, sprmList);
            }
            if (newPAP.GetFNoLnn() != oldPAP.GetFNoLnn())
            {
                size += SprmUtils.AddSprm((short)0x240C, newPAP.GetFNoLnn(), null, sprmList);
            }
            if (newPAP.GetFSideBySide() != oldPAP.GetFSideBySide())
            {
                size += SprmUtils.AddSprm((short)0x2404, newPAP.GetFSideBySide(), null, sprmList);
            }
            if (newPAP.GetFNoAutoHyph() != oldPAP.GetFNoAutoHyph())
            {
                size += SprmUtils.AddSprm((short)0x242A, newPAP.GetFNoAutoHyph(), null, sprmList);
            }
            if (newPAP.GetFWidowControl() != oldPAP.GetFWidowControl())
            {
                size += SprmUtils.AddSprm((short)0x2431, newPAP.GetFWidowControl(), null, sprmList);
            }
            if (newPAP.GetItbdMac() != oldPAP.GetItbdMac() ||
                !Arrays.Equals(newPAP.GetRgdxaTab(), oldPAP.GetRgdxaTab()) ||
                !Arrays.Equals(newPAP.GetRgtbd(), oldPAP.GetRgtbd()))
            {
                /** @todo revisit this */
                //      byte[] oldTabArray = oldPAP.GetRgdxaTab();
                //      byte[] newTabArray = newPAP.GetRgdxaTab();
                //      byte[] newTabDescriptors = newPAP.GetRgtbd();
                //      byte[] varParam = new byte[2 + oldTabArray.Length + newTabArray.Length +
                //                                 newTabDescriptors.Length];
                //      varParam[0] = (byte)(oldTabArray.Length/2);
                //      int offset = 1;
                //      Array.Copy(oldTabArray, 0, varParam, offset, oldTabArray.Length);
                //      offset += oldTabArray.Length;
                //      varParam[offset] = (byte)(newTabArray.Length/2);
                //      offset += 1;
                //      Array.Copy(newTabArray, 0, varParam, offset, newTabArray.Length);
                //      offset += newTabArray.Length;
                //      Array.Copy(newTabDescriptors, 0, varParam, offset, newTabDescriptors.Length);
                //
                //      size += SprmUtils.AddSprm((short)0xC60D, 0, varParam, sprmList);
            }
            if (newPAP.GetDxaRight() != oldPAP.GetDxaRight())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x840E), newPAP.GetDxaRight(), null, sprmList);
            }
            if (newPAP.GetDxaLeft() != oldPAP.GetDxaLeft())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x840F), newPAP.GetDxaLeft(), null, sprmList);
            }
            if (newPAP.GetDxaLeft1() != oldPAP.GetDxaLeft1())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x8411), newPAP.GetDxaLeft1(), null, sprmList);
            }
            if (!newPAP.GetLspd().Equals(oldPAP.GetLspd()))
            {
                byte[] buf = new byte[4];
                newPAP.GetLspd().Serialize(buf, 0);

                size += SprmUtils.AddSprm((short)0x6412, LittleEndian.GetInt(buf), null, sprmList);
            }
            if (newPAP.GetDyaBefore() != oldPAP.GetDyaBefore())
            {
                size += SprmUtils.AddSprm(unchecked((short)0xA413), newPAP.GetDyaBefore(), null, sprmList);
            }
            if (newPAP.GetDyaAfter() != oldPAP.GetDyaAfter())
            {
                size += SprmUtils.AddSprm(unchecked((short)0xA414), newPAP.GetDyaAfter(), null, sprmList);
            }
            if (newPAP.GetDyaBefore() != oldPAP.GetDyaBefore())
            {
                size += SprmUtils.AddSprm((short)0x2404, newPAP.GetDyaBefore(), null, sprmList);
            }
            if (newPAP.GetFKinsoku() != oldPAP.GetFKinsoku())
            {
                size += SprmUtils.AddSprm((short)0x2433, newPAP.GetDyaBefore(), null, sprmList);
            }
            if (newPAP.GetFWordWrap() != oldPAP.GetFWordWrap())
            {
                size += SprmUtils.AddSprm((short)0x2434, newPAP.GetFWordWrap(), null, sprmList);
            }
            if (newPAP.GetFOverflowPunct() != oldPAP.GetFOverflowPunct())
            {
                size += SprmUtils.AddSprm((short)0x2435, newPAP.GetFOverflowPunct(), null, sprmList);
            }
            if (newPAP.GetFTopLinePunct() != oldPAP.GetFTopLinePunct())
            {
                size += SprmUtils.AddSprm((short)0x2436, newPAP.GetFTopLinePunct(), null, sprmList);
            }
            if (newPAP.GetFAutoSpaceDE() != oldPAP.GetFAutoSpaceDE())
            {
                size += SprmUtils.AddSprm((short)0x2437, newPAP.GetFAutoSpaceDE(), null, sprmList);
            }
            if (newPAP.GetFAutoSpaceDN() != oldPAP.GetFAutoSpaceDN())
            {
                size += SprmUtils.AddSprm((short)0x2438, newPAP.GetFAutoSpaceDN(), null, sprmList);
            }
            if (newPAP.GetWAlignFont() != oldPAP.GetWAlignFont())
            {
                size += SprmUtils.AddSprm((short)0x4439, newPAP.GetWAlignFont(), null, sprmList);
            }
            if (newPAP.IsFBackward() != oldPAP.IsFBackward() ||
                newPAP.IsFVertical() != oldPAP.IsFVertical() ||
                newPAP.IsFRotateFont() != oldPAP.IsFRotateFont())
            {
                int val = 0;
                if (newPAP.IsFBackward())
                {
                    val |= 0x2;
                }
                if (newPAP.IsFVertical())
                {
                    val |= 0x1;
                }
                if (newPAP.IsFRotateFont())
                {
                    val |= 0x4;
                }
                size += SprmUtils.AddSprm((short)0x443A, val, null, sprmList);
            }
            if (!Arrays.Equals(newPAP.GetAnld(), oldPAP.GetAnld()))
            {
                size += SprmUtils.AddSprm(unchecked((short)0xC63E), 0, newPAP.GetAnld(), sprmList);
            }
            if (newPAP.GetFInTable() != oldPAP.GetFInTable())
            {
                size += SprmUtils.AddSprm((short)0x2416, newPAP.GetFInTable(), null, sprmList);
            }
            if (newPAP.GetFTtp() != oldPAP.GetFTtp())
            {
                size += SprmUtils.AddSprm((short)0x2417, newPAP.GetFTtp(), null, sprmList);
            }
            if (newPAP.GetWr() != oldPAP.GetWr())
            {
                size += SprmUtils.AddSprm((short)0x2423, newPAP.GetWr(), null, sprmList);
            }
            if (newPAP.GetFLocked() != oldPAP.GetFLocked())
            {
                size += SprmUtils.AddSprm((short)0x2430, newPAP.GetFLocked(), null, sprmList);
            }
            if (newPAP.GetDxaAbs() != oldPAP.GetDxaAbs())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x8418), newPAP.GetDxaAbs(), null, sprmList);
            }
            if (newPAP.GetDyaAbs() != oldPAP.GetDyaAbs())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x8419), newPAP.GetDyaAbs(), null, sprmList);
            }
            if (newPAP.GetDxaWidth() != oldPAP.GetDxaWidth())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x841A), newPAP.GetDxaWidth(), null, sprmList);
            }
            if (!newPAP.GetBrcTop().Equals(oldPAP.GetBrcTop()))
            {
                int brc = newPAP.GetBrcTop().ToInt();
                size += SprmUtils.AddSprm((short)0x6424, brc, null, sprmList);
            }
            if (!newPAP.GetBrcLeft().Equals(oldPAP.GetBrcLeft()))
            {
                int brc = newPAP.GetBrcLeft().ToInt();
                size += SprmUtils.AddSprm((short)0x6425, brc, null, sprmList);
            }
            if (!newPAP.GetBrcBottom().Equals(oldPAP.GetBrcBottom()))
            {
                int brc = newPAP.GetBrcBottom().ToInt();
                size += SprmUtils.AddSprm((short)0x6426, brc, null, sprmList);
            }
            if (!newPAP.GetBrcRight().Equals(oldPAP.GetBrcRight()))
            {
                int brc = newPAP.GetBrcRight().ToInt();
                size += SprmUtils.AddSprm((short)0x6427, brc, null, sprmList);
            }
            if (newPAP.GetBrcBar().Equals(oldPAP.GetBrcBar()))
            {
                int brc = newPAP.GetBrcBar().ToInt();
                size += SprmUtils.AddSprm((short)0x6428, brc, null, sprmList);
            }
            if (newPAP.GetDxaFromText() != oldPAP.GetDxaFromText())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x842F), newPAP.GetDxaFromText(), null, sprmList);
            }
            if (newPAP.GetDyaFromText() != oldPAP.GetDyaFromText())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x842E), newPAP.GetDyaFromText(), null, sprmList);
            }
            if (newPAP.GetDyaHeight() != oldPAP.GetDyaHeight() ||
                newPAP.GetFMinHeight() != oldPAP.GetFMinHeight())
            {
                short val = (short)newPAP.GetDyaHeight();
                if (newPAP.GetFMinHeight() > 0)
                {
                    val |= unchecked((short)0x8000);
                }
                size += SprmUtils.AddSprm((short)0x442B, val, null, sprmList);
            }
            if (newPAP.GetShd() != null && !newPAP.GetShd().Equals(oldPAP.GetShd()))
            {
                size += SprmUtils.AddSprm((short)0x442D, newPAP.GetShd().ToShort(), null, sprmList);
            }
            if (newPAP.GetDcs() != null && !newPAP.GetDcs().Equals(oldPAP.GetDcs()))
            {
                size += SprmUtils.AddSprm((short)0x442C, newPAP.GetDcs().ToShort(), null, sprmList);
            }
            if (newPAP.GetLvl() != oldPAP.GetLvl())
            {
                size += SprmUtils.AddSprm((short)0x2640, newPAP.GetLvl(), null, sprmList);
            }
            if (newPAP.GetFNumRMIns() != oldPAP.GetFNumRMIns())
            {
                size += SprmUtils.AddSprm((short)0x2443, newPAP.GetFNumRMIns(), null, sprmList);
            }
            if (newPAP.GetFPropRMark() != oldPAP.GetFPropRMark() ||
                newPAP.GetIbstPropRMark() != oldPAP.GetIbstPropRMark() ||
                !newPAP.GetDttmPropRMark().Equals(oldPAP.GetDttmPropRMark()))
            {
                byte[] buf = new byte[7];
                buf[0] = (byte)newPAP.GetFPropRMark();
                LittleEndian.PutShort(buf, 1, (short)newPAP.GetIbstPropRMark());
                newPAP.GetDttmPropRMark().Serialize(buf, 3);
                size += SprmUtils.AddSprm(unchecked((short)0xC63F), 0, buf, sprmList);
            }
            if (!Arrays.Equals(newPAP.GetNumrm(), oldPAP.GetNumrm()))
            {
                size += SprmUtils.AddSprm(unchecked((short)0xC645), 0, newPAP.GetNumrm(), sprmList);
            }

            if (newPAP.GetTableLevel() != oldPAP.GetTableLevel())
            {
                size += SprmUtils.AddSprm((short)0x6649, newPAP.GetTableLevel(), null, sprmList);
            }

            if (newPAP.GetEmbeddedCellMark() != oldPAP.GetEmbeddedCellMark())
            {
                size += SprmUtils.AddSprm((short)0x244b, newPAP.GetEmbeddedCellMark(), null, sprmList);
            }

            if (newPAP.GetFTtpEmbedded() != oldPAP.GetFTtpEmbedded())
            {
                size += SprmUtils.AddSprm((short)0x244c, newPAP.GetFTtpEmbedded(), null, sprmList);
            }

            return SprmUtils.GetGrpprl(sprmList, size);

        }
    }
}