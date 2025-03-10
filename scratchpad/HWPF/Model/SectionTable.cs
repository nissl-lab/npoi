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

using System.Collections.Generic;
using NPOI.Util;
using NPOI.HWPF.Model.IO;
using System;
namespace NPOI.HWPF.Model
{



    /**
     * @author Ryan Ackley
     */
    public class SectionTable
    {
        private static int SED_SIZE = 12;

        protected List<SEPX> _sections = new List<SEPX>();
        protected List<TextPiece> _text;

        /** So we can know if things are unicode or not */
        private TextPieceTable tpt;

        public SectionTable()
        {
        }


        public SectionTable(byte[] documentStream, byte[] tableStream, int OffSet,
                            int size, int fcMin,
                            TextPieceTable tpt, int mainLength)
        {
            PlexOfCps sedPlex = new PlexOfCps(tableStream, OffSet, size, SED_SIZE);
            this.tpt = tpt;
            this._text = tpt.TextPieces;

            int length = sedPlex.Length;

            for (int x = 0; x < length; x++)
            {
                GenericPropertyNode node = sedPlex.GetProperty(x);
                SectionDescriptor sed = new SectionDescriptor(node.Bytes, 0);

                int fileOffset = sed.GetFc();
                //int startAt = CPtoFC(node.Start);
                //int endAt = CPtoFC(node.End);
                int startAt = node.Start;
                int endAt = node.End;

                // check for the optimization
                if (fileOffset == unchecked((int)0xffffffff))
                {
                    _sections.Add(new SEPX(sed, startAt, endAt, Array.Empty<byte>()));
                }
                else
                {
                    // The first short at the offset is the size of the grpprl.
                    int sepxSize = LittleEndian.GetShort(documentStream, fileOffset);
                    byte[] buf = new byte[sepxSize];
                    fileOffset += LittleEndianConsts.SHORT_SIZE;
                    Array.Copy(documentStream, fileOffset, buf, 0, buf.Length);
                    _sections.Add(new SEPX(sed, startAt, endAt, buf));
                }
            }

            // Some files seem to lie about their unicode status, which
            //  is very very pesky. Try to work around these, but this
            //  is Getting on for black magic...
            int mainEndsAt = mainLength;
            bool matchAt = false;
            bool matchHalf = false;
            for (int i = 0; i < _sections.Count; i++)
            {
                SEPX s = _sections[i];
                if (s.End == mainEndsAt)
                {
                    matchAt = true;
                }
                else if (s.End == mainEndsAt || s.End == mainEndsAt - 1)
                {
                    matchHalf = true;
                }
            }
            if (!matchAt && matchHalf)
            {
                //System.err.println("Your document seemed to be mostly unicode, but the section defInition was in bytes! Trying anyway, but things may well go wrong!");
                for (int i = 0; i < _sections.Count; i++)
                {
                    SEPX s = _sections[i];
                    GenericPropertyNode node = sedPlex.GetProperty(i);

                    int startAt = node.Start;
                    int endAt = node.End;
                    s.Start = (startAt);
                    s.End = (endAt);
                }
            }
        }

        public void AdjustForInsert(int listIndex, int Length)
        {
            int size = _sections.Count;
            SEPX sepx = _sections[listIndex];
            sepx.End = (sepx.End + Length);

            for (int x = listIndex + 1; x < size; x++)
            {
                sepx = _sections[x];
                sepx.Start = (sepx.Start + Length);
                sepx.End = (sepx.End + Length);
            }
        }

        // goss version of CPtoFC - this takes into account non-contiguous textpieces
        // that we have come across in real world documents. Tests against the example
        // code in HWPFDocument show no variation to Ryan's version of the code in
        // normal use, but this version works with our non-contiguous test case.
        // So far unable to get this test case to be written out as well due to
        // other issues. - piers
        private int CPtoFC(int CP)
        {
            TextPiece TP = null;

            for (int i = _text.Count - 1; i > -1; i--)
            {
                TP = _text[i];

                if (CP >= TP.GetCP()) break;
            }
            int FC = TP.PieceDescriptor.FilePosition;
            int offset = CP - TP.GetCP();
            if (TP.IsUnicode)
            {
                offset = offset * 2;
            }
            FC = FC + offset;
            return FC;
        }

        public List<SEPX> GetSections()
        {
            return _sections;
        }

        public void WriteTo(HWPFFileSystem sys, int fcMin)
        {
            HWPFStream docStream = sys.GetStream("WordDocument");
            HWPFStream tableStream = sys.GetStream("1Table");

            int offset = docStream.Offset;
            int len = _sections.Count;
            PlexOfCps plex = new PlexOfCps(SED_SIZE);

            for (int x = 0; x < len; x++)
            {
                SEPX sepx = _sections[x];
                byte[] grpprl = sepx.GetGrpprl();

                // write the sepx to the document stream. starts with a 2 byte size
                // followed by the grpprl
                byte[] shortBuf = new byte[2];
                LittleEndian.PutShort(shortBuf, (short)grpprl.Length);

                docStream.Write(shortBuf);
                docStream.Write(grpprl);

                // set the fc in the section descriptor
                SectionDescriptor sed = sepx.GetSectionDescriptor();
                sed.SetFc(offset);

                // add the section descriptor bytes to the PlexOfCps.


                // original line -
                //GenericPropertyNode property = new GenericPropertyNode(sepx.Start, sepx.End, sed.ToArray());

                // Line using Ryan's FCtoCP() conversion method -
                // unable to observe any effect on our testcases when using this code - piers
                GenericPropertyNode property = new GenericPropertyNode(tpt.GetCharIndex(sepx.StartBytes), tpt.GetCharIndex(sepx.EndBytes), sed.ToArray());


                plex.AddProperty(property);

                offset = docStream.Offset;
            }
            tableStream.Write(plex.ToByteArray());
        }
    }


}