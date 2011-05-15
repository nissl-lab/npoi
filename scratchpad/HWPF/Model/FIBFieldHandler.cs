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

    public class FIBFieldHandler
    {
        public static int STSHFORIG = 0;
        public static int STSHF = 1;
        public static int PLCFFNDREF = 2;
        public static int PLCFFNDTXT = 3;
        public static int PLCFANDREF = 4;
        public static int PLCFANDTXT = 5;
        public static int PLCFSED = 6;
        public static int PLCFPAD = 7;
        public static int PLCFPHE = 8;
        public static int STTBGLSY = 9;
        public static int PLCFGLSY = 10;
        public static int PLCFHDD = 11;
        public static int PLCFBTECHPX = 12;
        public static int PLCFBTEPAPX = 13;
        public static int PLCFSEA = 14;
        public static int STTBFFFN = 15;
        public static int PLCFFLDMOM = 16;
        public static int PLCFFLDHDR = 17;
        public static int PLCFFLDFTN = 18;
        public static int PLCFFLDATN = 19;
        public static int PLCFFLDMCR = 20;
        public static int STTBFBKMK = 21;
        public static int PLCFBKF = 22;
        public static int PLCFBKL = 23;
        public static int CMDS = 24;
        public static int PLCMCR = 25;
        public static int STTBFMCR = 26;
        public static int PRDRVR = 27;
        public static int PRENVPORT = 28;
        public static int PRENVLAND = 29;
        public static int WSS = 30;
        public static int DOP = 31;
        public static int STTBFASSOC = 32;
        public static int CLX = 33;
        public static int PLCFPGDFTN = 34;
        public static int AUTOSAVESOURCE = 35;
        public static int GRPXSTATNOWNERS = 36;//validated
        public static int STTBFATNBKMK = 37;
        public static int PLCFDOAMOM = 38;
        public static int PLCDOAHDR = 39;
        public static int PLCSPAMOM = 40;
        public static int PLCSPAHDR = 41;
        public static int PLCFATNBKF = 42;
        public static int PLCFATNBKL = 43;
        public static int PMS = 44;
        public static int FORMFLDSTTBS = 45;
        public static int PLCFENDREF = 46;
        public static int PLCFENDTXT = 47;
        public static int PLCFFLDEDN = 48;
        public static int PLCFPGDEDN = 49;
        public static int DGGINFO = 50;
        public static int STTBFRMARK = 51;
        public static int STTBCAPTION = 52;
        public static int STTBAUTOCAPTION = 53;
        public static int PLCFWKB = 54;
        public static int PLCFSPL = 55;
        public static int PLCFTXBXTXT = 56;
        public static int PLCFFLDTXBX = 57;//validated
        public static int PLCFHDRTXBXTXT = 58;
        public static int PLCFFLDHDRTXBX = 59;
        public static int STWUSER = 60;
        public static int STTBTTMBD = 61;
        public static int UNUSED = 62;
        public static int PGDMOTHER = 63;
        public static int BKDMOTHER = 64;
        public static int PGDFTN = 65;
        public static int BKDFTN = 66;
        public static int PGDEDN = 67;
        public static int BKDEDN = 68;
        public static int STTBFINTFLD = 69;
        public static int ROUTESLIP = 70;
        public static int STTBSAVEDBY = 71;
        public static int STTBFNM = 72;
        public static int PLCFLST = 73;
        public static int PLFLFO = 74;
        public static int PLCFTXBXBKD = 75;//validated
        public static int PLCFTXBXHDRBKD = 76;
        public static int DOCUNDO = 77;
        public static int RGBUSE = 78;
        public static int USP = 79;
        public static int USKF = 80;
        public static int PLCUPCRGBUSE = 81;
        public static int PLCUPCUSP = 82;
        public static int STTBGLSYSTYLE = 83;
        public static int PLGOSL = 84;
        public static int PLCOCX = 85;
        public static int PLCFBTELVC = 86;
        public static int MODIFIED = 87;
        public static int PLCFLVC = 88;
        public static int PLCASUMY = 89;
        public static int PLCFGRAM = 90;
        public static int STTBLISTNAMES = 91;
        public static int STTBFUSSR = 92;

        //private static POILogger log = POILogFactory.GetLogger(FIBFieldHandler.class);

        private static int FIELD_SIZE = LittleEndianConstants.INT_SIZE * 2;

        private Hashtable _unknownMap = new Hashtable();
        private int[] _fields;


        public FIBFieldHandler(byte[] mainStream, int offset, byte[] tableStream,
                               ArrayList offsetList, bool areKnown)
        {
            int numFields = LittleEndian.GetShort(mainStream, offset);
            offset += LittleEndianConstants.SHORT_SIZE;
            _fields = new int[numFields * 2];

            for (int x = 0; x < numFields; x++)
            {
                int fieldOffset = (x * FIELD_SIZE) + offset;
                int dsOffset = LittleEndian.GetInt(mainStream, fieldOffset);
                fieldOffset += LittleEndianConstants.INT_SIZE;
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
            return (_fields.Length * LittleEndianConstants.INT_SIZE) + LittleEndianConstants.SHORT_SIZE;
        }

        internal void WriteTo(byte[] mainStream, int offset, HWPFStream tableStream)
        {
            int length = _fields.Length / 2;
            LittleEndian.PutShort(mainStream, offset, (short)length);
            offset += LittleEndianConstants.SHORT_SIZE;

            for (int x = 0; x < length; x++)
            {
                UnhandledDataStructure ds = (UnhandledDataStructure)_unknownMap[x];
                if (ds != null)
                {
                    LittleEndian.PutInt(mainStream, offset, tableStream.Offset);
                    offset += LittleEndianConstants.INT_SIZE;
                    byte[] buf = ds.GetBuf();
                    tableStream.Write(buf);
                    LittleEndian.PutInt(mainStream, offset, buf.Length);
                    offset += LittleEndianConstants.INT_SIZE;
                }
                else
                {
                    LittleEndian.PutInt(mainStream, offset, _fields[x * 2]);
                    offset += LittleEndianConstants.INT_SIZE;
                    LittleEndian.PutInt(mainStream, offset, _fields[(x * 2) + 1]);
                    offset += LittleEndianConstants.INT_SIZE;
                }
            }
        }
    }

}
