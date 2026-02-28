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

    using NPOI.HWPF.UserModel;
    using NPOI.Util;

    public class TableSprmUncompressor
      : SprmUncompressor
    {
        public TableSprmUncompressor()
        {
        }

        public static TableProperties UncompressTAP(byte[] grpprl,
                                                        int Offset)
        {
            TableProperties newProperties = new TableProperties();

            SprmIterator sprmIt = new SprmIterator(grpprl, Offset);

            while (sprmIt.HasNext())
            {
                SprmOperation sprm = sprmIt.Next();

                //TAPXs are actually PAPXs so we have to make sure we are only trying to
                //uncompress the right type of sprm.
                if (sprm.Type == SprmOperation.TYPE_TAP)
                {
                    UncompressTAPOperation(newProperties, sprm);
                }
            }

            return newProperties;
        }
        public static TableProperties UncompressTAP(SprmBuffer sprmBuffer)
        {
            TableProperties tableProperties;

            SprmOperation sprmOperation = sprmBuffer.FindSprm(unchecked((short)0xd608));
            if (sprmOperation != null)
            {
                byte[] grpprl = sprmOperation.Grpprl;
                int offset = sprmOperation.GrpprlOffset;
                short itcMac = grpprl[offset];
                tableProperties = new TableProperties(itcMac);
            }
            else
            {
                //logger.log(POILogger.WARN,
                //        "Some table rows didn't specify number of columns in SPRMs");
                tableProperties = new TableProperties((short)1);
            }

            for (SprmIterator iterator = sprmBuffer.Iterator(); iterator.HasNext(); )
            {
                SprmOperation sprm = iterator.Next();

                /*
                 * TAPXs are actually PAPXs so we have to make sure we are only
                 * trying to uncompress the right type of sprm.
                 */
                if (sprm.Type == SprmOperation.TYPE_TAP)
                {
                    try
                    {
                        UncompressTAPOperation(tableProperties, sprm);
                    }
                    catch ( IndexOutOfRangeException ex)
                    {
                        //logger.log(POILogger.ERROR, "Unable to apply ", sprm,
                        //        ": ", ex, ex);
                    }
                }
            }
            return tableProperties;
        }
        /**
         * Used to uncompress a table property. Performs an operation defined
         * by a sprm stored in a tapx.
         *
         * @param newTAP The TableProperties object to perform the operation on.
         * @param operand The operand that defines this operation.
         * @param param The parameter for this operation.
         * @param varParam Variable length parameter for this operation.
         */
        static void UncompressTAPOperation(TableProperties newTAP, SprmOperation sprm)
        {
            switch (sprm.Operation)
            {
                case 0:
                    newTAP.SetJc((short)sprm.Operand);
                    break;
                case 0x01:
                    {
                        short[] rgdxaCenter = newTAP.GetRgdxaCenter();
                        short itcMac = newTAP.GetItcMac();
                        int adjust = sprm.Operand - (rgdxaCenter[0] + newTAP.GetDxaGapHalf());
                        for (int x = 0; x < itcMac; x++)
                        {
                            rgdxaCenter[x] += (short)adjust;
                        }
                        break;
                    }
                case 0x02:
                    {
                        short[] rgdxaCenter = newTAP.GetRgdxaCenter();
                        if (rgdxaCenter != null)
                        {
                            int adjust = newTAP.GetDxaGapHalf() - sprm.Operand;
                            rgdxaCenter[0] += (short)adjust;
                        }
                        newTAP.SetDxaGapHalf(sprm.Operand);
                        break;
                    }
                case 0x03:
                    newTAP.SetFCantSplit(GetFlag(sprm.Operand));
                    break;
                case 0x04:
                    newTAP.SetFTableHeader(GetFlag(sprm.Operand));
                    break;
                case 0x05:
                    {
                        byte[] buf = sprm.Grpprl;
                        int offset = sprm.GrpprlOffset;
                        newTAP.SetBrcTop(new BorderCode(buf, offset));
                        offset += BorderCode.SIZE;
                        newTAP.SetBrcLeft(new BorderCode(buf, offset));
                        offset += BorderCode.SIZE;
                        newTAP.SetBrcBottom(new BorderCode(buf, offset));
                        offset += BorderCode.SIZE;
                        newTAP.SetBrcRight(new BorderCode(buf, offset));
                        offset += BorderCode.SIZE;
                        newTAP.SetBrcHorizontal(new BorderCode(buf, offset));
                        offset += BorderCode.SIZE;
                        newTAP.SetBrcVertical(new BorderCode(buf, offset));
                        break;
                    }
                case 0x06:

                    //obsolete, used in word 1.x
                    break;
                case 0x07:
                    newTAP.SetDyaRowHeight(sprm.Operand);
                    break;
                case 0x08:
                    {
                        byte[] grpprl = sprm.Grpprl;
                        int offset = sprm.GrpprlOffset;
                        short itcMac = grpprl[offset];
                        short[] rgdxaCenter = new short[itcMac + 1];
                        TableCellDescriptor[] rgtc = new TableCellDescriptor[itcMac];
                        //I use varParam[0] and newTAP._itcMac interchangably
                        newTAP.SetItcMac(itcMac);
                        newTAP.SetRgdxaCenter(rgdxaCenter);
                        newTAP.SetRgtc(rgtc);

                        // get the rgdxaCenters
                        for (int x = 0; x < itcMac; x++)
                        {
                            rgdxaCenter[x] = LittleEndian.GetShort(grpprl, offset + (1 + (x * 2)));
                        }

                        // only try to get the TC entries if they exist...
                        int endOfSprm = offset + sprm.Size - 6; // -2 bytes for sprm - 2 for size short - 2 to correct Offsets being 0 based
                        int startOfTCs = offset + (1 + (itcMac + 1) * 2);

                        bool hasTCs = startOfTCs < endOfSprm;

                        for (int x = 0; x < itcMac; x++)
                        {
                            // Sometimes, the grpprl does not contain data at every Offset. I have no idea why this happens.
                            if (hasTCs && offset + (1 + ((itcMac + 1) * 2) + (x * 20)) < grpprl.Length)
                                rgtc[x] = TableCellDescriptor.ConvertBytesToTC(grpprl,
                                   offset + (1 + ((itcMac + 1) * 2) + (x * 20)));
                            else
                                rgtc[x] = new TableCellDescriptor();
                        }

                        rgdxaCenter[itcMac] = LittleEndian.GetShort(grpprl, offset + (1 + (itcMac * 2)));
                        break;
                    }
                case 0x09:

                    /** @todo handle cell shading*/
                    break;
                case 0x0a:

                    /** @todo handle word defined table styles*/
                    break;
                case 0x20:
                    //      {
                    //        TableCellDescriptor[] rgtc = newTAP.GetRgtc();
                    //
                    //        for (int x = varParam[0]; x < varParam[1]; x++)
                    //        {
                    //
                    //          if ((varParam[2] & 0x08) > 0)
                    //          {
                    //            short[] brcRight = rgtc[x].GetBrcRight ();
                    //            brcRight[0] = LittleEndian.Getshort (varParam, 6);
                    //            brcRight[1] = LittleEndian.Getshort (varParam, 8);
                    //          }
                    //          else if ((varParam[2] & 0x04) > 0)
                    //          {
                    //            short[] brcBottom = rgtc[x].GetBrcBottom ();
                    //            brcBottom[0] = LittleEndian.Getshort (varParam, 6);
                    //            brcBottom[1] = LittleEndian.Getshort (varParam, 8);
                    //          }
                    //          else if ((varParam[2] & 0x02) > 0)
                    //          {
                    //            short[] brcLeft = rgtc[x].GetBrcLeft ();
                    //            brcLeft[0] = LittleEndian.Getshort (varParam, 6);
                    //            brcLeft[1] = LittleEndian.Getshort (varParam, 8);
                    //          }
                    //          else if ((varParam[2] & 0x01) > 0)
                    //          {
                    //            short[] brcTop = rgtc[x].GetBrcTop ();
                    //            brcTop[0] = LittleEndian.Getshort (varParam, 6);
                    //            brcTop[1] = LittleEndian.Getshort (varParam, 8);
                    //          }
                    //        }
                    //        break;
                    //      }
                    break;
                case 0x21:
                    {
                        int param = sprm.Operand;
                        int index = (int)(param & 0xff000000) >> 24;
                        int count = (param & 0x00ff0000) >> 16;
                        int width = (param & 0x0000ffff);
                        int itcMac = newTAP.GetItcMac();

                        short[] rgdxaCenter = new short[itcMac + count + 1];
                        TableCellDescriptor[] rgtc = new TableCellDescriptor[itcMac + count];
                        if (index >= itcMac)
                        {
                            index = itcMac;
                            Array.Copy(newTAP.GetRgdxaCenter(), 0, rgdxaCenter, 0,
                                             itcMac + 1);
                            Array.Copy(newTAP.GetRgtc(), 0, rgtc, 0, itcMac);
                        }
                        else
                        {
                            //copy rgdxaCenter
                            Array.Copy(newTAP.GetRgdxaCenter(), 0, rgdxaCenter, 0,
                                             index + 1);
                            Array.Copy(newTAP.GetRgdxaCenter(), index + 1, rgdxaCenter,
                                             index + count, itcMac - (index));
                            //copy rgtc
                            Array.Copy(newTAP.GetRgtc(), 0, rgtc, 0, index);
                            Array.Copy(newTAP.GetRgtc(), index, rgtc, index + count,
                                             itcMac - index);
                        }

                        for (int x = index; x < index + count; x++)
                        {
                            rgtc[x] = new TableCellDescriptor();
                            rgdxaCenter[x] = (short)(rgdxaCenter[x - 1] + width);
                        }
                        rgdxaCenter[index +
                          count] = (short)(rgdxaCenter[(index + count) - 1] + width);
                        break;
                    }
                /**@todo handle table sprms from complex files*/
                case 0x22:
                case 0x23:
                case 0x24:
                case 0x25:
                case 0x26:
                case 0x27:
                case 0x28:
                case 0x29:
                case 0x2a:
                case 0x2b:
                case 0x2c:
                    break;
                default:
                    break;
            }
        }



    }
}
