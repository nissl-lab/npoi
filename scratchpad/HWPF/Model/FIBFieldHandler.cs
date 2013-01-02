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
    using System.Collections;
    using NPOI.Util;
    using NPOI.HWPF.Model.IO;
    using System.Collections.Generic;

    public class FIBFieldHandler
    {
        public const int STSHFORIG = 0;
        public const int STSHF = 1;
        public const int PLCFFNDREF = 2;
        public const int PLCFFNDTXT = 3;
        public const int PLCFANDREF = 4;
        public const int PLCFANDTXT = 5;
        public const int PLCFSED = 6;
        public const int PLCFPAD = 7;
        public const int PLCFPHE = 8;
        public const int STTBGLSY = 9;
        public const int PLCFGLSY = 10;
        public const int PLCFHDD = 11;
        public const int PLCFBTECHPX = 12;
        public const int PLCFBTEPAPX = 13;
        public const int PLCFSEA = 14;
        public const int STTBFFFN = 15;
        public const int PLCFFLDMOM = 16;
        public const int PLCFFLDHDR = 17;
        public const int PLCFFLDFTN = 18;
        public const int PLCFFLDATN = 19;
        public const int PLCFFLDMCR = 20;
        public const int STTBFBKMK = 21;
        public const int PLCFBKF = 22;
        public const int PLCFBKL = 23;
        public const int CMDS = 24;
        public const int PLCMCR = 25;
        public const int STTBFMCR = 26;
        public const int PRDRVR = 27;
        public const int PRENVPORT = 28;
        public const int PRENVLAND = 29;
        public const int WSS = 30;
        public const int DOP = 31;
        public const int STTBFASSOC = 32;
        public const int CLX = 33;
        public const int PLCFPGDFTN = 34;
        public const int AUTOSAVESOURCE = 35;
        public const int GRPXSTATNOWNERS = 36;//validated
        public const int STTBFATNBKMK = 37;
        public const int PLCFDOAMOM = 38;
        public const int PLCDOAHDR = 39;
        public const int PLCSPAMOM = 40;
        public const int PLCSPAHDR = 41;
        public const int PLCFATNBKF = 42;
        public const int PLCFATNBKL = 43;
        public const int PMS = 44;
        public const int FORMFLDSTTBS = 45;
        public const int PLCFENDREF = 46;
        public const int PLCFENDTXT = 47;
        public const int PLCFFLDEDN = 48;
        public const int PLCFPGDEDN = 49;
        public const int DGGINFO = 50;
        public const int STTBFRMARK = 51;
        public const int STTBCAPTION = 52;
        public const int STTBAUTOCAPTION = 53;
        public const int PLCFWKB = 54;
        public const int PLCFSPL = 55;
        public const int PLCFTXBXTXT = 56;
        public const int PLCFFLDTXBX = 57;//validated
        public const int PLCFHDRTXBXTXT = 58;
        public const int PLCFFLDHDRTXBX = 59;
        public const int STWUSER = 60;
        public const int STTBTTMBD = 61;
        public const int UNUSED = 62;
        public const int PGDMOTHER = 63;
        public const int BKDMOTHER = 64;
        public const int PGDFTN = 65;
        public const int BKDFTN = 66;
        public const int PGDEDN = 67;
        public const int BKDEDN = 68;
        public const int STTBFINTFLD = 69;
        public const int ROUTESLIP = 70;
        public const int STTBSAVEDBY = 71;
        public const int STTBFNM = 72;
        public const int PLCFLST = 73;
        public const int PLFLFO = 74;
        public const int PLCFTXBXBKD = 75;//validated
        public const int PLCFTXBXHDRBKD = 76;
        public const int DOCUNDO = 77;
        public const int RGBUSE = 78;
        public const int USP = 79;
        public const int USKF = 80;
        public const int PLCUPCRGBUSE = 81;
        public const int PLCUPCUSP = 82;
        public const int STTBGLSYSTYLE = 83;
        public const int PLGOSL = 84;
        public const int PLCOCX = 85;
        public const int PLCFBTELVC = 86;
        public const int MODIFIED = 87;
        public const int PLCFLVC = 88;
        public const int PLCASUMY = 89;
        public const int PLCFGRAM = 90;
        public const int STTBLISTNAMES = 91;
        public const int STTBFUSSR = 92;

        //private static POILogger log = POILogFactory.GetLogger(FIBFieldHandler.class);

        private static int FIELD_SIZE = LittleEndianConsts.INT_SIZE * 2;

        private Hashtable _unknownMap = new Hashtable();
        private int[] _fields;


        public FIBFieldHandler(byte[] mainStream, int offset, byte[] tableStream,
                               List<int> offsetList, bool areKnown)
        {
            int numFields = LittleEndian.GetShort(mainStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _fields = new int[numFields * 2];

            for (int x = 0; x < numFields; x++)
            {
                int fieldOffset = (x * FIELD_SIZE) + offset;
                int dsOffset = LittleEndian.GetInt(mainStream, fieldOffset);
                fieldOffset += LittleEndianConsts.INT_SIZE;
                int dsSize = LittleEndian.GetInt(mainStream, fieldOffset);

                if (offsetList.Contains(x) ^ areKnown)
                {
                    if (dsSize > 0)
                    {
                        if (dsOffset + dsSize > tableStream.Length)
                        {
                            //log.log(POILogger.WARN, "Unhandled data structure points to outside the buffer. " +
                            //                        "offset = " + dsOffset + ", length = " + dsSize +
                            //                        ", buffer length = " + tableStream.Length);
                        }
                        else
                        {
                            UnhandledDataStructure unhandled = new UnhandledDataStructure(
                              tableStream, dsOffset, dsSize);
                            _unknownMap.Add(x, unhandled);
                        }
                    }
                }
                _fields[x * 2] = dsOffset;
                _fields[(x * 2) + 1] = dsSize;
            }
        }

        public void ClearFields()
        {
            Arrays.Fill(_fields, 0);
        }

        public int GetFieldOffset(int field)
        {
            return _fields[field * 2];
        }

        public int GetFieldSize(int field)
        {
            return _fields[(field * 2) + 1];
        }

        public void SetFieldOffset(int field, int offset)
        {
            _fields[field * 2] = offset;
        }

        public void SetFieldSize(int field, int size)
        {
            _fields[(field * 2) + 1] = size;
        }

        public int SizeInBytes()
        {
            return (_fields.Length * LittleEndianConsts.INT_SIZE) + LittleEndianConsts.SHORT_SIZE;
        }

        internal void WriteTo(byte[] mainStream, int offset, HWPFStream tableStream)
        {
            int length = _fields.Length / 2;
            LittleEndian.PutShort(mainStream, offset, (short)length);
            offset += LittleEndianConsts.SHORT_SIZE;

            for (int x = 0; x < length; x++)
            {
                UnhandledDataStructure ds = (UnhandledDataStructure)_unknownMap[x];
                if (ds != null)
                {
                    LittleEndian.PutInt(mainStream, offset, tableStream.Offset);
                    offset += LittleEndianConsts.INT_SIZE;
                    byte[] buf = ds.GetBuf();
                    tableStream.Write(buf);
                    LittleEndian.PutInt(mainStream, offset, buf.Length);
                    offset += LittleEndianConsts.INT_SIZE;
                }
                else
                {
                    LittleEndian.PutInt(mainStream, offset, _fields[x * 2]);
                    offset += LittleEndianConsts.INT_SIZE;
                    LittleEndian.PutInt(mainStream, offset, _fields[(x * 2) + 1]);
                    offset += LittleEndianConsts.INT_SIZE;
                }
            }
        }
    }

}
