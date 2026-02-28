
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
    using NPOI.Util;


    public class SprmUtils
    {
        public SprmUtils()
        {
        }

        public static byte[] shortArrayToByteArray(short[] convert)
        {
            byte[] buf = new byte[convert.Length * LittleEndianConsts.SHORT_SIZE];

            for (int x = 0; x < convert.Length; x++)
            {
                LittleEndian.PutShort(buf, x * LittleEndianConsts.SHORT_SIZE, convert[x]);
            }

            return buf;
        }

        public static int AddSpecialSprm(short instruction, byte[] varParam, IList list)
        {
            byte[] sprm = new byte[varParam.Length + 4];
            System.Array.Copy(varParam, 0, sprm, 4, varParam.Length);
            LittleEndian.PutShort(sprm, instruction);
            LittleEndian.PutShort(sprm, 2, (short)(varParam.Length + 1));
            list.Add(sprm);
            return sprm.Length;
        }


        public static int AddSprm(short instruction, bool param,
            IList list)
        {
            return AddSprm(instruction, param ? 1 : 0, null, list);
        }

        public static int AddSprm(short instruction, int param, byte[] varParam, IList list)
        {
            int type = (instruction & 0xe000) >> 13;

            byte[] sprm = null;
            switch (type)
            {
                case 0:
                case 1:
                    sprm = new byte[3];
                    sprm[2] = (byte)param;
                    break;
                case 2:
                    sprm = new byte[4];
                    LittleEndian.PutShort(sprm, 2, (short)param);
                    break;
                case 3:
                    sprm = new byte[6];
                    LittleEndian.PutInt(sprm, 2, param);
                    break;
                case 4:
                case 5:
                    sprm = new byte[4];
                    LittleEndian.PutShort(sprm, 2, (short)param);
                    break;
                case 6:
                    int varLength=0;
                    if (varParam != null)
                    {
                        varLength=varParam.Length;
                    }
                    sprm = new byte[3 + varLength];
                    sprm[2] = (byte)varLength;
                    if (varLength != 0)
                    {
                        System.Array.Copy(varParam, 0, sprm, 3, varLength);
                    }
                    break;
                case 7:
                    sprm = new byte[5];
                    // this Is a three byte int so it has to be handled special
                    byte[] temp = new byte[4];
                    LittleEndian.PutInt(temp, 0, param);
                    System.Array.Copy(temp, 0, sprm, 2, 3);
                    break;
                default:
                    //should never happen
                    break;
            }
            LittleEndian.PutShort(sprm, 0, instruction);
            list.Add(sprm);
            return sprm.Length;
        }

        public static byte[] GetGrpprl(IList sprmList, int size)
        {
            // spit out the grpprl
            byte[] grpprl = new byte[size];
            int listSize = sprmList.Count - 1;
            int index = 0;
            for (; listSize >= 0; listSize--)
            {
                byte[] sprm = (byte[])sprmList[0];
                 sprmList.RemoveAt(0);
                System.Array.Copy(sprm, 0, grpprl, index, sprm.Length);
                index += sprm.Length;
            }

            return grpprl;

        }

        public static int ConvertBrcToInt(short[] brc)
        {
            byte[] buf = new byte[4];
            LittleEndian.PutShort(buf, brc[0]);
            LittleEndian.PutShort(buf, LittleEndianConsts.SHORT_SIZE, brc[1]);
            return LittleEndian.GetInt(buf);
        }
    }
}