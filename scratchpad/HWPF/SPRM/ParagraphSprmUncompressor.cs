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

namespace NPOI.HWPF.SPRM
{
    using System;
    using System.Collections;
    using NPOI.HWPF.UserModel;
    using NPOI.Util;
    using NPOI.Util.Collections;


    public class ParagraphSprmUncompressor
      : SprmUncompressor
    {
        public ParagraphSprmUncompressor()
        {
        }

        public static ParagraphProperties UncompressPAP(ParagraphProperties parent,
                                                        byte[] grpprl,
                                                        int Offset)
        {
            ParagraphProperties newProperties = null;
            newProperties = (ParagraphProperties)parent.Clone();

            SprmIterator sprmIt = new SprmIterator(grpprl, Offset);

            while (sprmIt.HasNext())
            {
                SprmOperation sprm = sprmIt.Next();

                // PAPXs can contain table sprms if the paragraph marks the end of a
                // table row
                if (sprm.Type == SprmOperation.TYPE_PAP)
                {
                    UncompressPAPOperation(newProperties, sprm);
                }
            }

            return newProperties;
        }

        /**
         * Performs an operation on a ParagraphProperties object. Used to uncompress
         * from a papx.
         *
         * @param newPAP The ParagraphProperties object to perform the operation on.
         * @param operand The operand that defines the operation.
         * @param param The operation's parameter.
         * @param varParam The operation's variable length parameter.
         * @param grpprl The original papx.
         * @param offset The current offset in the papx.
         * @param spra A part of the sprm that defined this operation.
         */
        static void UncompressPAPOperation(ParagraphProperties newPAP, SprmOperation sprm)
        {
            switch (sprm.Operation)
            {
                case 0:
                    newPAP.SetIstd(sprm.Operand);
                    break;
                case 0x1:

                    // Used only for piece table grpprl's not for PAPX
                    //        int istdFirst = LittleEndian.Getshort (varParam, 2);
                    //        int istdLast = LittleEndian.Getshort (varParam, 4);
                    //        if ((newPAP.GetIstd () > istdFirst) || (newPAP.GetIstd () <= istdLast))
                    //        {
                    //          permuteIstd (newPAP, varParam, opSize);
                    //        }
                    break;
                case 0x2:
                    if (newPAP.GetIstd() <= 9 || newPAP.GetIstd() >= 1)
                    {
                        byte paramTmp = (byte)sprm.Operand;
                        newPAP.SetIstd(newPAP.GetIstd() + paramTmp);
                        newPAP.SetLvl((byte)(newPAP.GetLvl() + paramTmp));

                        if (((paramTmp >> 7) & 0x01) == 1)
                        {
                            newPAP.SetIstd(Math.Max(newPAP.GetIstd(), 1));
                        }
                        else
                        {
                            newPAP.SetIstd(Math.Min(newPAP.GetIstd(), 9));
                        }

                    }
                    break;
                case 0x3:
                    // Physical justification of the paragraph
                    newPAP.SetJc((byte)sprm.Operand);
                    break;
                case 0x4:
                    newPAP.SetFSideBySide(sprm.Operand!=0);
                    break;
                case 0x5:
                    newPAP.SetFKeep(sprm.Operand!=0);
                    break;
                case 0x6:
                    newPAP.SetFKeepFollow(sprm.Operand!=0);
                    break;
                case 0x7:
                    newPAP.SetFPageBreakBefore(sprm.Operand!=0);
                    break;
                case 0x8:
                    newPAP.SetBrcl((byte)sprm.Operand);
                    break;
                case 0x9:
                    newPAP.SetBrcp((byte)sprm.Operand);
                    break;
                case 0xa:
                    newPAP.SetIlvl((byte)sprm.Operand);
                    break;
                case 0xb:
                    newPAP.SetIlfo(sprm.Operand);
                    break;
                case 0xc:
                    newPAP.SetFNoLnn(sprm.Operand!=0);
                    break;
                case 0xd:
                    /**handle tabs . variable parameter. seperate Processing needed*/
                    handleTabs(newPAP, sprm);
                    break;
                case 0xe:
                    newPAP.SetDxaRight(sprm.Operand);
                    break;
                case 0xf:
                    newPAP.SetDxaLeft(sprm.Operand);
                    break;
                case 0x10:

                    // sprmPNest is only stored in grpprls linked to a piece table.
                    newPAP.SetDxaLeft(newPAP.GetDxaLeft() + sprm.Operand);
                    newPAP.SetDxaLeft(Math.Max(0, newPAP.GetDxaLeft()));
                    break;
                case 0x11:
                    newPAP.SetDxaLeft1(sprm.Operand);
                    break;
                case 0x12:
                    newPAP.SetLspd(new LineSpacingDescriptor(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x13:
                    newPAP.SetDyaBefore(sprm.Operand);
                    break;
                case 0x14:
                    newPAP.SetDyaAfter(sprm.Operand);
                    break;
                case 0x15:
                    // fast saved only
                    //ApplySprmPChgTabs (newPAP, varParam, opSize);
                    break;
                case 0x16:
                    newPAP.SetFInTable(sprm.Operand!=0);
                    break;
                case 0x17:
                    newPAP.SetFTtp(sprm.Operand!=0);
                    break;
                case 0x18:
                    newPAP.SetDxaAbs(sprm.Operand);
                    break;
                case 0x19:
                    newPAP.SetDyaAbs(sprm.Operand);
                    break;
                case 0x1a:
                    newPAP.SetDxaWidth(sprm.Operand);
                    break;
                case 0x1b:
                    byte param = (byte)sprm.Operand;
                    /** @todo handle paragraph postioning*/
                    byte pcVert = (byte)((param & 0x0c) >> 2);
                    byte pcHorz = (byte)(param & 0x03);
                    if (pcVert != 3)
                    {
                        newPAP.SetPcVert(pcVert);
                    }
                    if (pcHorz != 3)
                    {
                        newPAP.SetPcHorz(pcHorz);
                    }
                    break;

                // BrcXXX1 is older Version. Brc is used
                case 0x1c:

                    //newPAP.SetBrcTop1((short)param);
                    break;
                case 0x1d:

                    //newPAP.SetBrcLeft1((short)param);
                    break;
                case 0x1e:

                    //newPAP.SetBrcBottom1((short)param);
                    break;
                case 0x1f:

                    //newPAP.SetBrcRight1((short)param);
                    break;
                case 0x20:

                    //newPAP.SetBrcBetween1((short)param);
                    break;
                case 0x21:

                    //newPAP.SetBrcBar1((byte)param);
                    break;
                case 0x22:
                    newPAP.SetDxaFromText(sprm.Operand);
                    break;
                case 0x23:
                    newPAP.SetWr((byte)sprm.Operand);
                    break;
                case 0x24:
                    newPAP.SetBrcTop(new BorderCode(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x25:
                    newPAP.SetBrcLeft(new BorderCode(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x26:
                    newPAP.SetBrcBottom(new BorderCode(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x27:
                    newPAP.SetBrcRight(new BorderCode(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x28:
                    newPAP.SetBrcBetween(new BorderCode(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x29:
                    newPAP.SetBrcBar(new BorderCode(sprm.Grpprl, sprm.GrpprlOffset));
                    break;
                case 0x2a:
                    newPAP.SetFNoAutoHyph(sprm.Operand!=0);
                    break;
                case 0x2b:
                    newPAP.SetDyaHeight(sprm.Operand);
                    break;
                case 0x2c:
                    newPAP.SetDcs(new DropCapSpecifier((short)sprm.Operand));
                    break;
                case 0x2d:
                    newPAP.SetShd(new ShadingDescriptor((short)sprm.Operand));
                    break;
                case 0x2e:
                    newPAP.SetDyaFromText(sprm.Operand);
                    break;
                case 0x2f:
                    newPAP.SetDxaFromText(sprm.Operand);
                    break;
                case 0x30:
                    newPAP.SetFLocked(sprm.Operand!=0);
                    break;
                case 0x31:
                    newPAP.SetFWidowControl(sprm.Operand!=0);
                    break;
                case 0x32:

                    //undocumented
                    break;
                case 0x33:
                    newPAP.SetFKinsoku(sprm.Operand!=0);
                    break;
                case 0x34:
                    newPAP.SetFWordWrap(sprm.Operand!=0);
                    break;
                case 0x35:
                    newPAP.SetFOverflowPunct(sprm.Operand!=0);
                    break;
                case 0x36:
                    newPAP.SetFTopLinePunct(sprm.Operand!=0);
                    break;
                case 0x37:
                    newPAP.SetFAutoSpaceDE(sprm.Operand!=0);
                    break;
                case 0x38:
                    newPAP.SetFAutoSpaceDN(sprm.Operand!=0);
                    break;
                case 0x39:
                    newPAP.SetWAlignFont(sprm.Operand);
                    break;
                case 0x3a:
                    newPAP.SetFontAlign((short)sprm.Operand);
                    break;
                case 0x3b:

                    //obsolete
                    break;
                case 0x3e:
                    byte[] buf = new byte[sprm.Size - 3];
                    Array.Copy(buf, 0, sprm.Grpprl, sprm.GrpprlOffset,
                                     buf.Length);
                    newPAP.SetAnld(buf);
                    break;
                case 0x3f:
                    //don't really need this. spec is confusing regarding this
                    //sprm
                        byte[] varParam = sprm.Grpprl;
                        int offset = sprm.GrpprlOffset;
                        newPAP.SetFPropRMark(varParam[offset]!=0);
                        newPAP.SetIbstPropRMark(LittleEndian.GetShort(varParam, offset + 1));
                        newPAP.SetDttmPropRMark(new DateAndTime(varParam, offset + 3));

                    break;
                case 0x40:
                    // This condition commented out, as Word seems to set outline levels even for 
                    //  paragraph with other styles than Heading 1..9, even though specification 
                    //  does not say so. See bug 49820 for discussion.
                    //if (newPAP.GetIstd () < 1 && newPAP.GetIstd () > 9)
                    //{
                        newPAP.SetLvl((byte)sprm.Operand);
                    //}
                    break;
                case 0x41:

                    // undocumented
                    break;
                case 0x43:

                    //pap.fNumRMIns
                    newPAP.SetFNumRMIns(sprm.Operand!=0);
                    break;
                case 0x44:

                    //undocumented
                    break;
                case 0x45:
                    if (sprm.SizeCode == 6)
                    {
                        byte[] buf1 = new byte[sprm.Size - 3];
                        Array.Copy(buf1, 0, sprm.Grpprl, sprm.GrpprlOffset, buf1.Length);
                        newPAP.SetNumrm(buf1);
                    }
                    else
                    {
                        /**@todo handle large PAPX from data stream*/
                    }
                    break;

                case 0x47:
                    newPAP.SetFUsePgsuSettings(sprm.Operand!=0);
                    break;
                case 0x48:
                    newPAP.SetFAdjustRight(sprm.Operand!=0);
                    break;
                case 0x49:
                    // sprmPItap -- 0x6649
                    newPAP.SetItap(sprm.Operand);
                    break;
                case 0x4a:
                    // sprmPDtap -- 0x664a
                    newPAP.SetItap((byte)(newPAP.GetItap() + sprm.Operand));
                    break;
                case 0x4b:
                    // sprmPFInnerTableCell -- 0x244b
                    newPAP.SetFInnerTableCell(sprm.Operand!=0);
                    break;
                case 0x4c:
                    // sprmPFInnerTtp -- 0x244c
                    newPAP.SetFTtpEmbedded(sprm.Operand!=0);
                    break;
                case 0x61:
                    // sprmPJc 
                    newPAP.SetJustificationLogical((byte)sprm.Operand);
                    break;
                default:
                    break;
            }
        }

        private static void handleTabs(ParagraphProperties pap, SprmOperation sprm)
        {
            byte[] grpprl = sprm.Grpprl;
            int offset = sprm.GrpprlOffset;
            int delSize = grpprl[offset++];
            int[] tabPositions = pap.GetRgdxaTab();
            byte[] tabDescriptors = pap.GetRgtbd();

            Hashtable tabMap = new Hashtable();
            for (int x = 0; x < tabPositions.Length; x++)
            {
                tabMap.Add(tabPositions[x], tabDescriptors[x]);
            }

            for (int x = 0; x < delSize; x++)
            {
                tabMap.Remove(LittleEndian.GetShort(grpprl, offset));
                offset += LittleEndianConsts.SHORT_SIZE;
            }

            int addSize = grpprl[offset++];
            int start = offset;
            for (int x = 0; x < addSize; x++)
            {
                int key = LittleEndian.GetShort(grpprl, offset);
                Byte val = grpprl[start + ((LittleEndianConsts.SHORT_SIZE * addSize) + x)];
                tabMap.Add(key, val);
                offset += LittleEndianConsts.SHORT_SIZE;
            }

            tabPositions = new int[tabMap.Count];
            tabDescriptors = new byte[tabPositions.Length];
            ArrayList list = new ArrayList();

            IEnumerator keyIT = tabMap.Keys.GetEnumerator();
            while (keyIT.MoveNext())
            {
                list.Add(keyIT.Current);
            }
            list.Sort();

            for (int x = 0; x < tabPositions.Length; x++)
            {
                int key = (int)list[x];
                tabPositions[x] = key;
                tabDescriptors[x] = (byte)tabMap[key];
            }

            pap.SetRgdxaTab(tabPositions);
            pap.SetRgtbd(tabDescriptors);
        }

        //  private static void handleTabsAgain(ParagraphProperties pap, SprmOperation sprm)
        //  {
        //    byte[] grpprl = sprm.Grpprl;
        //    int offset = sprm.GrpprlOffset;
        //    int delSize = grpprl[Offset++];
        //    int[] tabPositions = pap.GetRgdxaTab();
        //    byte[] tabDescriptors = pap.GetRgtbd();
        //
        //    HashMap tabMap = new HashMap();
        //    for (int x = 0; x < tabPositions.Length; x++)
        //    {
        //      tabMap.Put(Int32.ValueOf(tabPositions[x]), Byte.ValueOf(tabDescriptors[x]));
        //    }
        //
        //    for (int x = 0; x < delSize; x++)
        //    {
        //      tabMap.Remove(Int32.valueOf(LittleEndian.getInt(grpprl, Offset)));
        //      offset += LittleEndian.INT_SIZE;;
        //    }
        //
        //    int AddSize = grpprl[Offset++];
        //    for (int x = 0; x < AddSize; x++)
        //    {
        //      Integer key = Int32.ValueOf(LittleEndian.getInt(grpprl, Offset));
        //      Byte val = Byte.ValueOf(grpprl[(LittleEndian.INT_SIZE * (AddSize - x)) + x]);
        //      tabMap.Put(key, val);
        //      offset += LittleEndian.INT_SIZE;
        //    }
        //
        //    tabPositions = new int[tabMap.Count];
        //    tabDescriptors = new byte[tabPositions.Length];
        //    ArrayList list = new ArrayList();
        //
        //    Iterator keyIT = tabMap.keySet().iterator();
        //    while (keyIT.hasNext())
        //    {
        //      list.Add(keyIT.next());
        //    }
        //    Collections.sort(list);
        //
        //    for (int x = 0; x < tabPositions.Length; x++)
        //    {
        //      Integer key = ((Integer)list.Get(x));
        //      tabPositions[x] = key.intValue();
        //      tabDescriptors[x] = ((Byte)tabMap.Get(key)).byteValue();
        //    }
        //
        //    pap.SetRgdxaTab(tabPositions);
        //    pap.SetRgtbd(tabDescriptors);
        //  }

    }

}