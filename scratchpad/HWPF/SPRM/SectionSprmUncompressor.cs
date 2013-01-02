
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

    public class SectionSprmUncompressor : SprmUncompressor
    {
        public SectionSprmUncompressor()
        {
        }
        public static SectionProperties UncompressSEP(byte[] grpprl, int offset)
        {
            SectionProperties newProperties = new SectionProperties();

            SprmIterator sprmIt = new SprmIterator(grpprl, offset);

            while (sprmIt.HasNext())
            {
                SprmOperation sprm = (SprmOperation)sprmIt.Next();
                UncompressSEPOperation(newProperties, sprm);
            }

            return newProperties;
        }

        /**
         * Used in decompression of a sepx. This performs an operation defined by
         * a single sprm.
         *
         * @param newSEP The SectionProperty to perfrom the operation on.
         * @param operand The operation to perform.
         * @param param The operation's parameter.
         * @param varParam The operation variable length parameter.
         */
        static void UncompressSEPOperation(SectionProperties newSEP, SprmOperation sprm)
        {
            switch (sprm.Operation)
            {
                case 0:
                    newSEP.SetCnsPgn((byte)sprm.Operand);
                    break;
                case 0x1:
                    newSEP.SetIHeadingPgn((byte)sprm.Operand);
                    break;
                case 0x2:
                    byte[] buf = new byte[sprm.Size - 3];
                    Array.Copy(sprm.Grpprl, sprm.GrpprlOffset, buf, 0, buf.Length);
                    newSEP.SetOlstAnm(buf);
                    break;
                case 0x3:
                    //not quite sure
                    break;
                case 0x4:
                    //not quite sure
                    break;
                case 0x5:
                    newSEP.SetFEvenlySpaced(GetFlag(sprm.Operand));
                    break;
                case 0x6:
                    newSEP.SetFUnlocked(GetFlag(sprm.Operand));
                    break;
                case 0x7:
                    newSEP.SetDmBinFirst((short)sprm.Operand);
                    break;
                case 0x8:
                    newSEP.SetDmBinOther((short)sprm.Operand);
                    break;
                case 0x9:
                    newSEP.SetBkc((byte)sprm.Operand);
                    break;
                case 0xa:
                    newSEP.SetFTitlePage(GetFlag(sprm.Operand));
                    break;
                case 0xb:
                    newSEP.SetCcolM1((short)sprm.Operand);
                    break;
                case 0xc:
                    newSEP.SetDxaColumns(sprm.Operand);
                    break;
                case 0xd:
                    newSEP.SetFAutoPgn(GetFlag(sprm.Operand));
                    break;
                case 0xe:
                    newSEP.SetNfcPgn((byte)sprm.Operand);
                    break;
                case 0xf:
                    newSEP.SetDyaPgn((short)sprm.Operand);
                    break;
                case 0x10:
                    newSEP.SetDxaPgn((short)sprm.Operand);
                    break;
                case 0x11:
                    newSEP.SetFPgnRestart(GetFlag(sprm.Operand));
                    break;
                case 0x12:
                    newSEP.SetFEndNote(GetFlag(sprm.Operand));
                    break;
                case 0x13:
                    newSEP.SetLnc((byte)sprm.Operand);
                    break;
                case 0x14:
                    newSEP.SetGrpfIhdt((byte)sprm.Operand);
                    break;
                case 0x15:
                    newSEP.SetNLnnMod((short)sprm.Operand);
                    break;
                case 0x16:
                    newSEP.SetDxaLnn(sprm.Operand);
                    break;
                case 0x17:
                    newSEP.SetDyaHdrTop(sprm.Operand);
                    break;
                case 0x18:
                    newSEP.SetDyaHdrBottom(sprm.Operand);
                    break;
                case 0x19:
                    newSEP.SetFLBetween(GetFlag(sprm.Operand));
                    break;
                case 0x1a:
                    newSEP.SetVjc((byte)sprm.Operand);
                    break;
                case 0x1b:
                    newSEP.SetLnnMin((short)sprm.Operand);
                    break;
                case 0x1c:
                    newSEP.SetPgnStart((short)sprm.Operand);
                    break;
                case 0x1d:
                    newSEP.SetDmOrientPage(sprm.Operand!=0);
                    break;
                case 0x1e:

                    //nothing
                    break;
                case 0x1f:
                    newSEP.SetXaPage(sprm.Operand);
                    break;
                case 0x20:
                    newSEP.SetYaPage(sprm.Operand);
                    break;
                case 0x21:
                    newSEP.SetDxaLeft(sprm.Operand);
                    break;
                case 0x22:
                    newSEP.SetDxaRight(sprm.Operand);
                    break;
                case 0x23:
                    newSEP.SetDyaTop(sprm.Operand);
                    break;
                case 0x24:
                    newSEP.SetDyaBottom(sprm.Operand);
                    break;
                case 0x25:
                    newSEP.SetDzaGutter(sprm.Operand);
                    break;
                case 0x26:
                    newSEP.SetDmPaperReq((short)sprm.Operand);
                    break;
                case 0x27:
                    newSEP.SetFPropMark(GetFlag(sprm.Operand));
                    break;
                case 0x28:
                    break;
                case 0x29:
                    break;
                case 0x2a:
                    break;
                case 0x2b:
                    newSEP.SetBrcTop(new BorderCode(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x2c:
                    newSEP.SetBrcLeft(new BorderCode(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x2d:
                    newSEP.SetBrcBottom(new BorderCode(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x2e:
                    newSEP.SetBrcRight(new BorderCode(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x2f:
                    newSEP.SetPgbProp(sprm.Operand);
                    break;
                case 0x30:
                    newSEP.SetDxtCharSpace(sprm.Operand);
                    break;
                case 0x31:
                    newSEP.SetDyaLinePitch(sprm.Operand);
                    break;
                case 0x33:
                    newSEP.SetWTextFlow((short)sprm.Operand);
                    break;
                default:
                    break;
            }

        }


    }
}