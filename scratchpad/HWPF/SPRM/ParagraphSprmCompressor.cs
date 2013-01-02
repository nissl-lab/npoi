
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

        public static byte[] CompressParagraphProperty(ParagraphProperties newPAP,
                                                       ParagraphProperties oldPAP)
        {
            // page numbers links to Word97-2007BinaryFileFormat(doc)Specification.pdf, accessible from microsoft.com 
            ArrayList sprmList = new ArrayList();
            int size = 0;

            // Page 50 of public specification begins
            if (newPAP.GetIstd() != oldPAP.GetIstd())
            {
                // sprmPIstd 
                size += SprmUtils.AddSprm((short)0x4600, newPAP.GetIstd(), null, sprmList);
            }
            if (newPAP.GetJc() != oldPAP.GetJc())
            {
                // sprmPJc80 
                size += SprmUtils.AddSprm((short)0x2403, newPAP.GetJc(), null, sprmList);
            }
            if (newPAP.GetFSideBySide() != oldPAP.GetFSideBySide())
            {
                // sprmPFSideBySide 
                size += SprmUtils.AddSprm((short)0x2404, Convert.ToBoolean(newPAP.GetFSideBySide()), sprmList);
            }
            if (newPAP.GetFKeep() != oldPAP.GetFKeep())
            {
                size += SprmUtils.AddSprm((short)0x2405, newPAP.GetFKeep(), sprmList);
            }
            if (newPAP.GetFKeepFollow() != oldPAP.GetFKeepFollow())
            {
                size += SprmUtils.AddSprm((short)0x2406, newPAP.GetFKeepFollow(), sprmList);
            }
            if (newPAP.GetFPageBreakBefore() != oldPAP.GetFPageBreakBefore())
            {
                size += SprmUtils.AddSprm((short)0x2407, newPAP.GetFPageBreakBefore(), sprmList);
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
                size += SprmUtils.AddSprm((short)0x240C, newPAP.GetFNoLnn(), sprmList);
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
            if (newPAP.GetDxaLeft() != oldPAP.GetDxaLeft())
            {
                // sprmPDxaLeft80 
                size += SprmUtils.AddSprm(unchecked((short)0x840F), newPAP.GetDxaLeft(), null, sprmList);
            }

            // Page 51 of public specification begins
            if (newPAP.GetDxaLeft1() != oldPAP.GetDxaLeft1())
            {
                // sprmPDxaLeft180 
                size += SprmUtils.AddSprm(unchecked((short)0x8411), newPAP.GetDxaLeft1(), null, sprmList);
            }
            if (newPAP.GetDxaRight() != oldPAP.GetDxaRight())
            {
                // sprmPDxaRight80  
                size += SprmUtils.AddSprm(unchecked((short)0x840E), newPAP.GetDxaRight(), null, sprmList);
            }
            if (newPAP.GetDxcLeft() != oldPAP.GetDxcLeft())
            {
                // sprmPDxcLeft
                size += SprmUtils.AddSprm((short)0x4456, newPAP.GetDxcLeft(), null, sprmList);
            }
            if (newPAP.GetDxcLeft1() != oldPAP.GetDxcLeft1())
            {
                // sprmPDxcLeft1
                size += SprmUtils.AddSprm((short)0x4457, newPAP.GetDxcLeft1(), null, sprmList);
            }
            if (newPAP.GetDxcRight() != oldPAP.GetDxcRight())
            {
                // sprmPDxcRight
                size += SprmUtils.AddSprm((short)0x4455, newPAP.GetDxcRight(), null, sprmList);
            }
            if (!newPAP.GetLspd().Equals(oldPAP.GetLspd()))
            {
                // sprmPDyaLine
                byte[] buf = new byte[4];
                newPAP.GetLspd().Serialize(buf, 0);

                size += SprmUtils.AddSprm((short)0x6412, LittleEndian.GetInt(buf), null, sprmList);
            }
            if (newPAP.GetDyaBefore() != oldPAP.GetDyaBefore())
            {
                // sprmPDyaBefore
                size += SprmUtils.AddSprm(unchecked((short)0xA413), newPAP.GetDyaBefore(), null, sprmList);
            }
            if (newPAP.GetDyaAfter() != oldPAP.GetDyaAfter())
            {
                // sprmPDyaAfter
                size += SprmUtils.AddSprm(unchecked((short)0xA414), newPAP.GetDyaAfter(), null, sprmList);
            }
            if (newPAP.GetFDyaBeforeAuto() != oldPAP.GetFDyaBeforeAuto())
            {
                // sprmPFDyaBeforeAuto
                size += SprmUtils.AddSprm((short)0x245B, newPAP.GetFDyaBeforeAuto(), sprmList);
            }
            if (newPAP.GetFDyaAfterAuto() != oldPAP.GetFDyaAfterAuto())
            {
                // sprmPFDyaAfterAuto
                size += SprmUtils.AddSprm((short)0x245C, newPAP.GetFDyaAfterAuto(), sprmList);
            }
            if (newPAP.GetFInTable() != oldPAP.GetFInTable())
            {
                // sprmPFInTable
                size += SprmUtils.AddSprm((short)0x2416, newPAP.GetFInTable(), sprmList);
            }
            if (newPAP.GetFTtp() != oldPAP.GetFTtp())
            {
                // sprmPFTtp
                size += SprmUtils.AddSprm((short)0x2417, newPAP.GetFTtp(), sprmList);
            }
            if (newPAP.GetDxaAbs() != oldPAP.GetDxaAbs())
            {
                // sprmPDxaAbs
                size += SprmUtils.AddSprm(unchecked((short)0x8418), newPAP.GetDxaAbs(), null, sprmList);
            }
            if (newPAP.GetDyaAbs() != oldPAP.GetDyaAbs())
            {
                // sprmPDyaAbs
                size += SprmUtils.AddSprm(unchecked((short)0x8419), newPAP.GetDyaAbs(), null, sprmList);
            }

            if (newPAP.GetDxaWidth() != oldPAP.GetDxaWidth())
            {
                // sprmPDxaWidth
                size += SprmUtils.AddSprm(unchecked((short)0x841A), newPAP.GetDxaWidth(), null, sprmList);
            }


            if (newPAP.GetWr() != oldPAP.GetWr())
            {
                size += SprmUtils.AddSprm((short)0x2423, newPAP.GetWr(), null, sprmList);
            }
            if (newPAP.GetBrcBar().Equals(oldPAP.GetBrcBar()))
            {
                // XXX: sprm code 0x6428 is sprmPBrcBetween80, but accessed property linked to sprmPBrcBar80 (0x6629)
                int brc = newPAP.GetBrcBar().ToInt();
                size += SprmUtils.AddSprm((short)0x6428, brc, null, sprmList);
            }
            if (!newPAP.GetBrcBottom().Equals(oldPAP.GetBrcBottom()))
            {
                // sprmPBrcBottom80  
                int brc = newPAP.GetBrcBottom().ToInt();
                size += SprmUtils.AddSprm((short)0x6426, brc, null, sprmList);
            }
            if (!newPAP.GetBrcLeft().Equals(oldPAP.GetBrcLeft()))
            {
                // sprmPBrcLeft80  
                int brc = newPAP.GetBrcLeft().ToInt();
                size += SprmUtils.AddSprm((short)0x6425, brc, null, sprmList);
            }
            // Page 53 of public specification begins
            if (!newPAP.GetBrcRight().Equals(oldPAP.GetBrcRight()))
            {
                // sprmPBrcRight80
                int brc = newPAP.GetBrcRight().ToInt();
                size += SprmUtils.AddSprm((short)0x6427, brc, null, sprmList);
            }
            if (!newPAP.GetBrcTop().Equals(oldPAP.GetBrcTop()))
            {
                // sprmPBrcTop80 
                int brc = newPAP.GetBrcTop().ToInt();
                size += SprmUtils.AddSprm((short)0x6424, brc, null, sprmList);
            }
            if (newPAP.GetFNoAutoHyph() != oldPAP.GetFNoAutoHyph())
            {
                size += SprmUtils.AddSprm((short)0x242A, newPAP.GetFNoAutoHyph(), sprmList);
            }
            if (newPAP.GetDyaHeight() != oldPAP.GetDyaHeight() ||
                 newPAP.GetFMinHeight() != oldPAP.GetFMinHeight())
            {
                // sprmPWHeightAbs
                short val = (short)newPAP.GetDyaHeight();
                if (newPAP.GetFMinHeight())
                {
                    val |= unchecked((short)0x8000);
                }
                size += SprmUtils.AddSprm((short)0x442B, val, null, sprmList);
            }
            if (newPAP.GetDcs() != null && !newPAP.GetDcs().Equals(oldPAP.GetDcs()))
            {
                // sprmPDcs 
                size += SprmUtils.AddSprm((short)0x442C, newPAP.GetDcs().ToShort(), null, sprmList);
            }
            if (newPAP.GetShd() != null && !newPAP.GetShd().Equals(oldPAP.GetShd()))
            {
                // sprmPShd80 
                size += SprmUtils.AddSprm((short)0x442D, newPAP.GetShd().ToShort(), null, sprmList);
            }
            if (newPAP.GetDyaFromText() != oldPAP.GetDyaFromText())
            {
                // sprmPDyaFromText
                size += SprmUtils.AddSprm(unchecked((short)0x842E), newPAP.GetDyaFromText(), null, sprmList);
            }
            if (newPAP.GetDxaFromText() != oldPAP.GetDxaFromText())
            {
                // sprmPDxaFromText
                size += SprmUtils.AddSprm(unchecked((short)0x842F), newPAP.GetDxaFromText(), null, sprmList);
            }
            if (newPAP.GetFLocked() != oldPAP.GetFLocked())
            {
                // sprmPFLocked
                size += SprmUtils.AddSprm((short)0x2430, newPAP.GetFLocked(), sprmList);
            }
            if (newPAP.GetFWidowControl() != oldPAP.GetFWidowControl())
            {
                // sprmPFWidowControl
                size += SprmUtils.AddSprm((short)0x2431, newPAP.GetFWidowControl(), sprmList);
            }
            if (newPAP.GetFKinsoku() != oldPAP.GetFKinsoku())
            {
                size += SprmUtils.AddSprm((short)0x2433, newPAP.GetDyaBefore(), null, sprmList);
            }
            if (newPAP.GetFWordWrap() != oldPAP.GetFWordWrap())
            {
                size += SprmUtils.AddSprm((short)0x2434, newPAP.GetFWordWrap(), sprmList);
            }
            if (newPAP.GetFOverflowPunct() != oldPAP.GetFOverflowPunct())
            {
                size += SprmUtils.AddSprm((short)0x2435, newPAP.GetFOverflowPunct(), sprmList);
            }
            if (newPAP.GetFTopLinePunct() != oldPAP.GetFTopLinePunct())
            {
                size += SprmUtils.AddSprm((short)0x2436, newPAP.GetFTopLinePunct(), sprmList);
            }
            if (newPAP.GetFAutoSpaceDE() != oldPAP.GetFAutoSpaceDE())
            {
                size += SprmUtils.AddSprm((short)0x2437, newPAP.GetFAutoSpaceDE(), sprmList);
            }
            if (newPAP.GetFAutoSpaceDN() != oldPAP.GetFAutoSpaceDN())
            {
                size += SprmUtils.AddSprm((short)0x2438, newPAP.GetFAutoSpaceDN(), sprmList);
            }
            if (newPAP.GetWAlignFont() != oldPAP.GetWAlignFont())
            {
                size += SprmUtils.AddSprm((short)0x4439, newPAP.GetWAlignFont(), null, sprmList);
            }
            // Page 54 of public specification begins
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
                // sprmPAnld80 
                size += SprmUtils.AddSprm(unchecked((short)0xC63E), 0, newPAP.GetAnld(), sprmList);
            }
            if (newPAP.GetFPropRMark() != oldPAP.GetFPropRMark() ||
                newPAP.GetIbstPropRMark() != oldPAP.GetIbstPropRMark() ||
                !newPAP.GetDttmPropRMark().Equals(oldPAP.GetDttmPropRMark()))
            {
                // sprmPPropRMark
                byte[] buf = new byte[7];
                buf[0] = (byte)(newPAP.GetFPropRMark() ? 1 : 0);
                LittleEndian.PutShort(buf, 1, (short)newPAP.GetIbstPropRMark());
                newPAP.GetDttmPropRMark().Serialize(buf, 3);
                size += SprmUtils.AddSprm(unchecked((short)0xC63F), 0, buf, sprmList);
            }
            if (newPAP.GetLvl() != oldPAP.GetLvl())
            {
                // sprmPOutLvl 
                size += SprmUtils.AddSprm((short)0x2640, newPAP.GetLvl(), null, sprmList);
            }
            if (newPAP.GetFBiDi() != oldPAP.GetFBiDi())
            {
                // sprmPFBiDi 
                size += SprmUtils.AddSprm((short)0x2441, newPAP.GetFBiDi(), sprmList);
            }
            if (newPAP.GetFNumRMIns() != oldPAP.GetFNumRMIns())
            {
                // sprmPFNumRMIns 
                size += SprmUtils.AddSprm((short)0x2443, newPAP.GetFNumRMIns(), sprmList);
            }

            if (!Arrays.Equals(newPAP.GetNumrm(), oldPAP.GetNumrm()))
            {
                // sprmPNumRM
                size += SprmUtils.AddSprm(unchecked((short)0xC645), 0, newPAP.GetNumrm(), sprmList);
            }

            if (newPAP.GetFInnerTableCell() != oldPAP.GetFInnerTableCell())
            {
                // sprmPFInnerTableCell
                size += SprmUtils.AddSprm((short)0x244b, newPAP.GetFInnerTableCell(), sprmList);
            }

            if (newPAP.GetFTtpEmbedded() != oldPAP.GetFTtpEmbedded())
            {
                // sprmPFInnerTtp 
                size += SprmUtils.AddSprm((short)0x244c, newPAP.GetFTtpEmbedded(), sprmList);
            }
            // Page 55 of public specification begins
            if (newPAP.GetItap() != oldPAP.GetItap())
            {
                // sprmPItap
                size += SprmUtils.AddSprm((short)0x6649, newPAP.GetItap(), null, sprmList);
            }

            return SprmUtils.GetGrpprl(sprmList, size);

        }
    }
}