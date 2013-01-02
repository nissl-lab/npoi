
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
    using NPOI.HWPF.Model.Types;


    public class SectionSprmCompressor
    {
        private static SectionProperties DEFAULT_SEP = new SectionProperties();
        public SectionSprmCompressor()
        {
        }
        public static byte[] CompressSectionProperty(SectionProperties newSEP)
        {
            int size = 0;
            ArrayList sprmList = new ArrayList();

            if (newSEP.GetCnsPgn() != DEFAULT_SEP.GetCnsPgn())
            {
                size += SprmUtils.AddSprm((short)0x3000, newSEP.GetCnsPgn(), null, sprmList);
            }
            if (newSEP.GetIHeadingPgn() != DEFAULT_SEP.GetIHeadingPgn())
            {
                size += SprmUtils.AddSprm((short)0x3001, newSEP.GetIHeadingPgn(), null, sprmList);
            }
            if (!Arrays.Equals(newSEP.GetOlstAnm(), DEFAULT_SEP.GetOlstAnm()))
            {
                size += SprmUtils.AddSprm(unchecked((short)0xD202), 0, newSEP.GetOlstAnm(), sprmList);
            }
            if (newSEP.GetFEvenlySpaced() != DEFAULT_SEP.GetFEvenlySpaced())
            {
                size += SprmUtils.AddSprm((short)0x3005, newSEP.GetFEvenlySpaced() ? 1 : 0, null, sprmList);
            }
            if (newSEP.GetFUnlocked() != DEFAULT_SEP.GetFUnlocked())
            {
                size += SprmUtils.AddSprm((short)0x3006, newSEP.GetFUnlocked() ? 1 : 0, null, sprmList);
            }
            if (newSEP.GetDmBinFirst() != DEFAULT_SEP.GetDmBinFirst())
            {
                size += SprmUtils.AddSprm((short)0x5007, newSEP.GetDmBinFirst(), null, sprmList);
            }
            if (newSEP.GetDmBinOther() != DEFAULT_SEP.GetDmBinOther())
            {
                size += SprmUtils.AddSprm((short)0x5008, newSEP.GetDmBinOther(), null, sprmList);
            }
            if (newSEP.GetBkc() != DEFAULT_SEP.GetBkc())
            {
                size += SprmUtils.AddSprm((short)0x3009, newSEP.GetBkc(), null, sprmList);
            }
            if (newSEP.GetFTitlePage() != DEFAULT_SEP.GetFTitlePage())
            {
                size += SprmUtils.AddSprm((short)0x300A, newSEP.GetFTitlePage() ? 1 : 0, null, sprmList);
            }
            if (newSEP.GetCcolM1() != DEFAULT_SEP.GetCcolM1())
            {
                size += SprmUtils.AddSprm((short)0x500B, newSEP.GetCcolM1(), null, sprmList);
            }
            if (newSEP.GetDxaColumns() != DEFAULT_SEP.GetDxaColumns())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x900C), newSEP.GetDxaColumns(), null, sprmList);
            }
            if (newSEP.GetFAutoPgn() != DEFAULT_SEP.GetFAutoPgn())
            {
                size += SprmUtils.AddSprm((short)0x300D, newSEP.GetFAutoPgn() ? 1 : 0, null, sprmList);
            }
            if (newSEP.GetNfcPgn() != DEFAULT_SEP.GetNfcPgn())
            {
                size += SprmUtils.AddSprm((short)0x300E, newSEP.GetNfcPgn(), null, sprmList);
            }
            if (newSEP.GetDyaPgn() != DEFAULT_SEP.GetDyaPgn())
            {
                size += SprmUtils.AddSprm(unchecked((short)0xB00F), newSEP.GetDyaPgn(), null, sprmList);
            }
            if (newSEP.GetDxaPgn() != DEFAULT_SEP.GetDxaPgn())
            {
                size += SprmUtils.AddSprm(unchecked((short)0xB010), newSEP.GetDxaPgn(), null, sprmList);
            }
            if (newSEP.GetFPgnRestart() != DEFAULT_SEP.GetFPgnRestart())
            {
                size += SprmUtils.AddSprm((short)0x3011, newSEP.GetFPgnRestart() ? 1 : 0, null, sprmList);
            }
            if (newSEP.GetFEndNote() != DEFAULT_SEP.GetFEndNote())
            {
                size += SprmUtils.AddSprm((short)0x3012, newSEP.GetFEndNote() ? 1 : 0, null, sprmList);
            }
            if (newSEP.GetLnc() != DEFAULT_SEP.GetLnc())
            {
                size += SprmUtils.AddSprm((short)0x3013, newSEP.GetLnc(), null, sprmList);
            }
            if (newSEP.GetGrpfIhdt() != DEFAULT_SEP.GetGrpfIhdt())
            {
                size += SprmUtils.AddSprm((short)0x3014, newSEP.GetGrpfIhdt(), null, sprmList);
            }
            if (newSEP.GetNLnnMod() != DEFAULT_SEP.GetNLnnMod())
            {
                size += SprmUtils.AddSprm((short)0x5015, newSEP.GetNLnnMod(), null, sprmList);
            }
            if (newSEP.GetDxaLnn() != DEFAULT_SEP.GetDxaLnn())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x9016), newSEP.GetDxaLnn(), null, sprmList);
            }
            if (newSEP.GetDyaHdrTop() != DEFAULT_SEP.GetDyaHdrTop())
            {
                size += SprmUtils.AddSprm(unchecked((short)0xB017), newSEP.GetDyaHdrTop(), null, sprmList);
            }
            if (newSEP.GetDyaHdrBottom() != DEFAULT_SEP.GetDyaHdrBottom())
            {
                size += SprmUtils.AddSprm(unchecked((short)0xB018), newSEP.GetDyaHdrBottom(), null, sprmList);
            }
            if (newSEP.GetFLBetween() != DEFAULT_SEP.GetFLBetween())
            {
                size += SprmUtils.AddSprm((short)0x3019, newSEP.GetFLBetween() ? 1 : 0, null, sprmList);
            }
            if (newSEP.GetVjc() != DEFAULT_SEP.GetVjc())
            {
                size += SprmUtils.AddSprm((short)0x301A, newSEP.GetVjc(), null, sprmList);
            }
            if (newSEP.GetLnnMin() != DEFAULT_SEP.GetLnnMin())
            {
                size += SprmUtils.AddSprm((short)0x501B, newSEP.GetLnnMin(), null, sprmList);
            }
            if (newSEP.GetPgnStart() != DEFAULT_SEP.GetPgnStart())
            {
                size += SprmUtils.AddSprm((short)0x501C, newSEP.GetPgnStart(), null, sprmList);
            }
            if (newSEP.GetDmOrientPage() != DEFAULT_SEP.GetDmOrientPage())
            {
                size += SprmUtils.AddSprm((short)0x301D, newSEP.GetDmOrientPage()?1:0, null, sprmList);
            }
            if (newSEP.GetXaPage() != DEFAULT_SEP.GetXaPage())
            {
                size += SprmUtils.AddSprm(unchecked((short)0xB01F), newSEP.GetXaPage(), null, sprmList);
            }
            if (newSEP.GetYaPage() != DEFAULT_SEP.GetYaPage())
            {
                size += SprmUtils.AddSprm(unchecked((short)0xB020), newSEP.GetYaPage(), null, sprmList);
            }
            if (newSEP.GetDxaLeft() != DEFAULT_SEP.GetDxaLeft())
            {
                size += SprmUtils.AddSprm(unchecked((short)0xB021), newSEP.GetDxaLeft(), null, sprmList);
            }
            if (newSEP.GetDxaRight() != DEFAULT_SEP.GetDxaRight())
            {
                size += SprmUtils.AddSprm(unchecked((short)0xB022), newSEP.GetDxaRight(), null, sprmList);
            }
            if (newSEP.GetDyaTop() != DEFAULT_SEP.GetDyaTop())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x9023), newSEP.GetDyaTop(), null, sprmList);
            }
            if (newSEP.GetDyaBottom() != DEFAULT_SEP.GetDyaBottom())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x9024), newSEP.GetDyaBottom(), null, sprmList);
            }
            if (newSEP.GetDzaGutter() != DEFAULT_SEP.GetDzaGutter())
            {
                size += SprmUtils.AddSprm(unchecked((short)0xB025), newSEP.GetDzaGutter(), null, sprmList);
            }
            if (newSEP.GetDmPaperReq() != DEFAULT_SEP.GetDmPaperReq())
            {
                size += SprmUtils.AddSprm((short)0x5026, newSEP.GetDmPaperReq(), null, sprmList);
            }
            if (newSEP.GetFPropMark() != DEFAULT_SEP.GetFPropMark() ||
                newSEP.GetIbstPropRMark() != DEFAULT_SEP.GetIbstPropRMark() ||
                !newSEP.GetDttmPropRMark().Equals(DEFAULT_SEP.GetDttmPropRMark()))
            {
                byte[] buf = new byte[7];
                buf[0] = (byte)(newSEP.GetFPropMark() ? 1 : 0);
                int offset = LittleEndianConsts.BYTE_SIZE;
                LittleEndian.PutShort(buf, (short)newSEP.GetIbstPropRMark());
                offset += LittleEndianConsts.SHORT_SIZE;
                newSEP.GetDttmPropRMark().Serialize(buf, offset);
                size += SprmUtils.AddSprm(unchecked((short)0xD227), -1, buf, sprmList);
            }
            if (!newSEP.GetBrcTop().Equals(DEFAULT_SEP.GetBrcTop()))
            {
                size += SprmUtils.AddSprm((short)0x702B, newSEP.GetBrcTop().ToInt(), null, sprmList);
            }
            if (!newSEP.GetBrcLeft().Equals(DEFAULT_SEP.GetBrcLeft()))
            {
                size += SprmUtils.AddSprm((short)0x702C, newSEP.GetBrcLeft().ToInt(), null, sprmList);
            }
            if (!newSEP.GetBrcBottom().Equals(DEFAULT_SEP.GetBrcBottom()))
            {
                size += SprmUtils.AddSprm((short)0x702D, newSEP.GetBrcBottom().ToInt(), null, sprmList);
            }
            if (!newSEP.GetBrcRight().Equals(DEFAULT_SEP.GetBrcRight()))
            {
                size += SprmUtils.AddSprm((short)0x702E, newSEP.GetBrcRight().ToInt(), null, sprmList);
            }
            if (newSEP.GetPgbProp() != DEFAULT_SEP.GetPgbProp())
            {
                size += SprmUtils.AddSprm((short)0x522F, newSEP.GetPgbProp(), null, sprmList);
            }
            if (newSEP.GetDxtCharSpace() != DEFAULT_SEP.GetDxtCharSpace())
            {
                size += SprmUtils.AddSprm((short)0x7030, newSEP.GetDxtCharSpace(), null, sprmList);
            }
            if (newSEP.GetDyaLinePitch() != DEFAULT_SEP.GetDyaLinePitch())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x9031), newSEP.GetDyaLinePitch(), null, sprmList);
            }
            if (newSEP.GetClm() != DEFAULT_SEP.GetClm())
            {
                size += SprmUtils.AddSprm((short)0x5032, newSEP.GetClm(), null, sprmList);
            }
            if (newSEP.GetWTextFlow() != DEFAULT_SEP.GetWTextFlow())
            {
                size += SprmUtils.AddSprm((short)0x5033, newSEP.GetWTextFlow(), null, sprmList);
            }

            return SprmUtils.GetGrpprl(sprmList, size);
        }
    }

}