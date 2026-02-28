
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

    public class CharacterSprmCompressor
    {
        public CharacterSprmCompressor()
        {
        }
        public static byte[] CompressCharacterProperty(CharacterProperties newCHP, CharacterProperties oldCHP)
        {
            ArrayList sprmList = new ArrayList();
            int size = 0;

            if (newCHP.IsFRMarkDel() != oldCHP.IsFRMarkDel())
            {
                int value = 0;
                if (newCHP.IsFRMarkDel())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0800, value, null, sprmList);
            }
            if (newCHP.IsFRMark() != oldCHP.IsFRMark())
            {
                int value = 0;
                if (newCHP.IsFRMark())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0801, value, null, sprmList);
            }
            if (newCHP.IsFFldVanish() != oldCHP.IsFFldVanish())
            {
                int value = 0;
                if (newCHP.IsFFldVanish())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0802, value, null, sprmList);
            }
            if (newCHP.IsFSpec() != oldCHP.IsFSpec() || newCHP.GetFcPic() != oldCHP.GetFcPic())
            {
                size += SprmUtils.AddSprm((short)0x6a03, newCHP.GetFcPic(), null, sprmList);
            }
            if (newCHP.GetIbstRMark() != oldCHP.GetIbstRMark())
            {
                size += SprmUtils.AddSprm((short)0x4804, newCHP.GetIbstRMark(), null, sprmList);
            }
            if (!newCHP.GetDttmRMark().Equals(oldCHP.GetDttmRMark()))
            {
                byte[] buf = new byte[4];
                newCHP.GetDttmRMark().Serialize(buf, 0);

                size += SprmUtils.AddSprm((short)0x6805, LittleEndian.GetInt(buf), null, sprmList);
            }
            if (newCHP.IsFData() != oldCHP.IsFData())
            {
                int value = 0;
                if (newCHP.IsFData())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0806, value, null, sprmList);
            }
            if (newCHP.IsFSpec() && newCHP.GetFtcSym() != 0)
            {
                byte[] varParam = new byte[4];
                LittleEndian.PutShort(varParam, 0, (short)newCHP.GetFtcSym());
                LittleEndian.PutShort(varParam, 2, (short)newCHP.GetXchSym());

                size += SprmUtils.AddSprm((short)0x6a09, 0, varParam, sprmList);
            }
            if (newCHP.IsFOle2() != newCHP.IsFOle2())
            {
                int value = 0;
                if (newCHP.IsFOle2())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x080a, value, null, sprmList);
            }
            if (newCHP.GetIcoHighlight() != oldCHP.GetIcoHighlight())
            {
                size += SprmUtils.AddSprm((short)0x2a0c, newCHP.GetIcoHighlight(), null, sprmList);
            }
            if (newCHP.GetFcObj() != oldCHP.GetFcObj())
            {
                size += SprmUtils.AddSprm((short)0x680e, newCHP.GetFcObj(), null, sprmList);
            }
            if (newCHP.GetIstd() != oldCHP.GetIstd())
            {
                size += SprmUtils.AddSprm((short)0x4a30, newCHP.GetIstd(), null, sprmList);
            }
            if (newCHP.IsFBold() != oldCHP.IsFBold())
            {
                int value = 0;
                if (newCHP.IsFBold())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0835, value, null, sprmList);
            }
            if (newCHP.IsFItalic() != oldCHP.IsFItalic())
            {
                int value = 0;
                if (newCHP.IsFItalic())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0836, value, null, sprmList);
            }
            if (newCHP.IsFStrike() != oldCHP.IsFStrike())
            {
                int value = 0;
                if (newCHP.IsFStrike())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0837, value, null, sprmList);
            }
            if (newCHP.IsFOutline() != oldCHP.IsFOutline())
            {
                int value = 0;
                if (newCHP.IsFOutline())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0838, value, null, sprmList);
            }
            if (newCHP.IsFShadow() != oldCHP.IsFShadow())
            {
                int value = 0;
                if (newCHP.IsFShadow())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0839, value, null, sprmList);
            }
            if (newCHP.IsFSmallCaps() != oldCHP.IsFSmallCaps())
            {
                int value = 0;
                if (newCHP.IsFSmallCaps())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x083a, value, null, sprmList);
            }
            if (newCHP.IsFCaps() != oldCHP.IsFCaps())
            {
                int value = 0;
                if (newCHP.IsFCaps())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x083b, value, null, sprmList);
            }
            if (newCHP.IsFVanish() != oldCHP.IsFVanish())
            {
                int value = 0;
                if (newCHP.IsFVanish())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x083c, value, null, sprmList);
            }
            if (newCHP.GetKul() != oldCHP.GetKul())
            {
                size += SprmUtils.AddSprm((short)0x2a3e, newCHP.GetKul(), null, sprmList);
            }
            if (newCHP.GetDxaSpace() != oldCHP.GetDxaSpace())
            {
                size += SprmUtils.AddSprm(unchecked((short)0x8840), newCHP.GetDxaSpace(), null, sprmList);
            }
            if (newCHP.GetIco() != oldCHP.GetIco())
            {
                size += SprmUtils.AddSprm((short)0x2a42, newCHP.GetIco(), null, sprmList);
            }
            if (newCHP.GetHps() != oldCHP.GetHps())
            {
                size += SprmUtils.AddSprm((short)0x4a43, newCHP.GetHps(), null, sprmList);
            }
            if (newCHP.GetHpsPos() != oldCHP.GetHpsPos())
            {
                size += SprmUtils.AddSprm((short)0x4845, newCHP.GetHpsPos(), null, sprmList);
            }
            if (newCHP.GetHpsKern() != oldCHP.GetHpsKern())
            {
                size += SprmUtils.AddSprm((short)0x484b, newCHP.GetHpsKern(), null, sprmList);
            }
            if (newCHP.GetYsr() != oldCHP.GetYsr())
            {
                size += SprmUtils.AddSprm((short)0x484e, newCHP.GetYsr(), null, sprmList);
            }
            if (newCHP.GetFtcAscii() != oldCHP.GetFtcAscii())
            {
                size += SprmUtils.AddSprm((short)0x4a4f, newCHP.GetFtcAscii(), null, sprmList);
            }
            if (newCHP.GetFtcFE() != oldCHP.GetFtcFE())
            {
                size += SprmUtils.AddSprm((short)0x4a50, newCHP.GetFtcFE(), null, sprmList);
            }
            if (newCHP.GetFtcOther() != oldCHP.GetFtcOther())
            {
                size += SprmUtils.AddSprm((short)0x4a51, newCHP.GetFtcOther(), null, sprmList);
            }

            if (newCHP.IsFDStrike() != oldCHP.IsFDStrike())
            {
                int value = 0;
                if (newCHP.IsFDStrike())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x2a53, value, null, sprmList);
            }
            if (newCHP.IsFImprint() != oldCHP.IsFImprint())
            {
                int value = 0;
                if (newCHP.IsFImprint())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0854, value, null, sprmList);
            }
            if (newCHP.IsFSpec() != oldCHP.IsFSpec())
            {
                int value = 0;
                if (newCHP.IsFSpec())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0855, value, null, sprmList);
            }
            if (newCHP.IsFObj() != oldCHP.IsFObj())
            {
                int value = 0;
                if (newCHP.IsFObj())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0856, value, null, sprmList);
            }
            if (newCHP.IsFEmboss() != oldCHP.IsFEmboss())
            {
                int value = 0;
                if (newCHP.IsFEmboss())
                {
                    value = 0x01;
                }
                size += SprmUtils.AddSprm((short)0x0858, value, null, sprmList);
            }
            if (newCHP.GetSfxtText() != oldCHP.GetSfxtText())
            {
                size += SprmUtils.AddSprm((short)0x2859, newCHP.GetSfxtText(), null, sprmList);
            }
            if (newCHP.GetIco24() != oldCHP.GetIco24())
            {
                if (newCHP.GetIco24() != -1) // don't Add a sprm if we're looking at an ico = Auto
                    size += SprmUtils.AddSprm((short)0x6870, newCHP.GetIco24(), null, sprmList);
            }

            return SprmUtils.GetGrpprl(sprmList, size);
        }



    }
}