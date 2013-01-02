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

using NPOI.HWPF.Model.Types;
using System.Collections;
using NPOI.HWPF.Model.IO;
using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
namespace NPOI.HWPF.Model
{

    /**
     * The File Information Block (FIB). Holds pointers
     *  to various bits of the file, and lots of flags which
     *  specify properties of the document.
     *
     * The parent class, {@link FIBAbstractType}, holds the
     *  first 32 bytes, which make up the FibBase.
     * The next part, the fibRgW / FibRgW97, is handled
     *  by {@link FIBshortHandler}.
     * The next part, the fibRgLw / The FibRgLw97, is
     *  handled by the {@link FIBLongHandler}.
     * Finally, the rest of the fields are handled by
     *  the {@link FIBFieldHandler}.
     *
     * @author  andy
     */
    public class FileInformationBlock : FIBAbstractType
    {


        FIBLongHandler _longHandler;
        FIBshortHandler _shortHandler;
        FIBFieldHandler _fieldHandler;

        /** Creates a new instance of FileInformationBlock */
        public FileInformationBlock(byte[] mainDocument)
        {
            FillFields(mainDocument, 0);
        }

        public void FillVariableFields(byte[] mainDocument, byte[] tableStream)
        {
            _shortHandler = new FIBshortHandler(mainDocument);
            _longHandler = new FIBLongHandler(mainDocument, FIBshortHandler.START + _shortHandler.SizeInBytes());

            List<int> knownFieldSet = new List<int>();
            knownFieldSet.Add(FIBFieldHandler.STSHF);
            knownFieldSet.Add(FIBFieldHandler.CLX);
            knownFieldSet.Add(FIBFieldHandler.DOP);
            knownFieldSet.Add(FIBFieldHandler.PLCFBTECHPX);
            knownFieldSet.Add(FIBFieldHandler.PLCFBTEPAPX);
            knownFieldSet.Add(FIBFieldHandler.PLCFSED);
            knownFieldSet.Add(FIBFieldHandler.PLCFLST);
            knownFieldSet.Add(FIBFieldHandler.PLFLFO);

        // field info
        foreach ( FieldsDocumentPart part in Enum.GetValues(typeof(FieldsDocumentPart)) )
            knownFieldSet.Add((int)part);

            // bookmarks
            knownFieldSet.Add(FIBFieldHandler.PLCFBKF);
            knownFieldSet.Add(FIBFieldHandler.PLCFBKL);
            knownFieldSet.Add(FIBFieldHandler.STTBFBKMK);


            // notes
            foreach (NoteType noteType in NoteType.Values)
            {
                knownFieldSet.Add(noteType
                        .GetFibDescriptorsFieldIndex());
                knownFieldSet.Add(noteType
                        .GetFibTextPositionsFieldIndex());
            }

            knownFieldSet.Add(FIBFieldHandler.STTBFFFN);
            knownFieldSet.Add(FIBFieldHandler.STTBFRMARK);
            knownFieldSet.Add(FIBFieldHandler.STTBSAVEDBY);
            knownFieldSet.Add(FIBFieldHandler.MODIFIED);


            _fieldHandler = new FIBFieldHandler(mainDocument,
                                                FIBshortHandler.START + _shortHandler.SizeInBytes() + _longHandler.SizeInBytes(),
                                                tableStream, knownFieldSet, true);
        }
        public override String ToString()
        {
            StringBuilder stringBuilder = new StringBuilder(base.ToString());
            stringBuilder.Append("[FIB2]\n");
            stringBuilder.Append("\tSubdocuments info:\n");
            foreach (SubdocumentType type in Enum.GetValues(typeof(SubdocumentType)))
            {
                stringBuilder.Append("\t\t");
                stringBuilder.Append(type);
                stringBuilder.Append(" has length of ");
                stringBuilder.Append(GetSubdocumentTextStreamLength(type));
                stringBuilder.Append(" char(s)\n");
            }
            stringBuilder.Append("\tFields PLCF info:\n");
            foreach (FieldsDocumentPart part in Enum.GetValues(typeof(FieldsDocumentPart)))
            {
                stringBuilder.Append("\t\t");
                stringBuilder.Append(part);
                stringBuilder.Append(": PLCF starts at ");
                stringBuilder.Append(GetFieldsPlcfOffset(part));
                stringBuilder.Append(" and have length of ");
                stringBuilder.Append(GetFieldsPlcfLength(part));
                stringBuilder.Append("\n");
            }
            stringBuilder.Append("\tNotes PLCF info:\n");
            foreach (NoteType noteType in NoteType.Values)
            {
                stringBuilder.Append("\t\t");
                stringBuilder.Append(noteType);
                stringBuilder.Append(": descriptions starts ");
                stringBuilder.Append(GetNotesDescriptorsOffset(noteType));
                stringBuilder.Append(" and have length of ");
                stringBuilder.Append(GetNotesDescriptorsSize(noteType));
                stringBuilder.Append(" bytes\n");
                stringBuilder.Append("\t\t");
                stringBuilder.Append(noteType);
                stringBuilder.Append(": text positions starts ");
                stringBuilder.Append(GetNotesTextPositionsOffset(noteType));
                stringBuilder.Append(" and have length of ");
                stringBuilder.Append(GetNotesTextPositionsSize(noteType));
                stringBuilder.Append(" bytes\n");
            }
            try
            {
                stringBuilder.Append("\t.Net reflection info:\n");
                foreach (MethodInfo method in typeof(FileInformationBlock).GetMethods())
                {
                    if (!method.Name.StartsWith("get")
                            || !method.IsPublic
                            || method.IsStatic
                            || method.GetParameters().Length > 0)
                        continue;
                    stringBuilder.Append("\t\t");
                    stringBuilder.Append(method.Name);
                    stringBuilder.Append(" => ");
                    stringBuilder.Append(method.Invoke(this, null));
                    stringBuilder.Append("\n");
                }
            }
            catch (Exception exc)
            {
                stringBuilder.Append("(exc: " + exc.Message + ")");
            }
            stringBuilder.Append("[/FIB2]\n");
            return stringBuilder.ToString();
        }

        public int GetFcDop()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.DOP);
        }

        public void SetFcDop(int fcDop)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.DOP, fcDop);
        }

        public int GetLcbDop()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.DOP);
        }

        public void SetLcbDop(int lcbDop)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.DOP, lcbDop);
        }

        public int GetFcStshf()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.STSHF);
        }

        public int GetLcbStshf()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.STSHF);
        }

        public void SetFcStshf(int fcStshf)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.STSHF, fcStshf);
        }

        public void SetLcbStshf(int lcbStshf)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.STSHF, lcbStshf);
        }

        public int GetFcClx()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.CLX);
        }

        public int GetLcbClx()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.CLX);
        }

        public void SetFcClx(int fcClx)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.CLX, fcClx);
        }

        public void SetLcbClx(int lcbClx)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.CLX, lcbClx);
        }

        public int GetFcPlcfbteChpx()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFBTECHPX);
        }

        public int GetLcbPlcfbteChpx()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFBTECHPX);
        }

        public void SetFcPlcfbteChpx(int fcPlcfBteChpx)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFBTECHPX, fcPlcfBteChpx);
        }

        public void SetLcbPlcfbteChpx(int lcbPlcfBteChpx)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFBTECHPX, lcbPlcfBteChpx);
        }

        public int GetFcPlcfbtePapx()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFBTEPAPX);
        }

        public int GetLcbPlcfbtePapx()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFBTEPAPX);
        }

        public void SetFcPlcfbtePapx(int fcPlcfBtePapx)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFBTEPAPX, fcPlcfBtePapx);
        }

        public void SetLcbPlcfbtePapx(int lcbPlcfBtePapx)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFBTEPAPX, lcbPlcfBtePapx);
        }

        public int GetFcPlcfsed()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFSED);
        }

        public int GetLcbPlcfsed()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFSED);
        }

        public void SetFcPlcfsed(int fcPlcfSed)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFSED, fcPlcfSed);
        }

        public void SetLcbPlcfsed(int lcbPlcfSed)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFSED, lcbPlcfSed);
        }

        public int GetFcPlcfLst()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFLST);
        }

        public int GetLcbPlcfLst()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFLST);
        }

        public void SetFcPlcfLst(int fcPlcfLst)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFLST, fcPlcfLst);
        }

        public void SetLcbPlcfLst(int lcbPlcfLst)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFLST, lcbPlcfLst);
        }

        public int GetFcPlfLfo()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLFLFO);
        }

        public int GetLcbPlfLfo()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLFLFO);
        }

        /**
 * @return Offset in table stream of the STTBF that records bookmark names
 *         in the main document
 */
        public int GetFcSttbfbkmk()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.STTBFBKMK);
        }

        public void SetFcSttbfbkmk(int offset)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.STTBFBKMK, offset);
        }

        /**
         * @return Count of bytes in Sttbfbkmk
         */
        public int GetLcbSttbfbkmk()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.STTBFBKMK);
        }

        public void SetLcbSttbfbkmk(int length)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.STTBFBKMK, length);
        }

        /**
         * @return Offset in table stream of the PLCF that records the beginning CP
         *         offsets of bookmarks in the main document. See BKF structure
         *         definition.
         */
        public int GetFcPlcfbkf()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFBKF);
        }

        public void SetFcPlcfbkf(int offset)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFBKF, offset);
        }

        /**
         * @return Count of bytes in Plcfbkf
         */
        public int GetLcbPlcfbkf()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFBKF);
        }

        public void SetLcbPlcfbkf(int length)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFBKF, length);
        }

        /**
         * @return Offset in table stream of the PLCF that records the ending CP
         *         offsets of bookmarks recorded in the main document. No structure
         *         is stored in this PLCF.
         */
        public int GetFcPlcfbkl()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFBKL);
        }

        public void SetFcPlcfbkl(int offset)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFBKL, offset);
        }

        /**
         * @return Count of bytes in Plcfbkl
         */
        public int GetLcbPlcfbkl()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFBKL);
        }

        public void SetLcbPlcfbkl(int length)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFBKL, length);
        }

        public void SetFcPlfLfo(int fcPlfLfo)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLFLFO, fcPlfLfo);
        }

        public void SetLcbPlfLfo(int lcbPlfLfo)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLFLFO, lcbPlfLfo);
        }

        public int GetFcSttbfffn()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.STTBFFFN);
        }

        public int GetLcbSttbfffn()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.STTBFFFN);
        }

        public void SetFcSttbfffn(int fcSttbFffn)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.STTBFFFN, fcSttbFffn);
        }

        public void SetLcbSttbfffn(int lcbSttbFffn)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.STTBFFFN, lcbSttbFffn);
        }

        public int GetFcSttbfRMark()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.STTBFRMARK);
        }

        public int GetLcbSttbfRMark()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.STTBFRMARK);
        }

        public void SetFcSttbfRMark(int fcSttbfRMark)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.STTBFRMARK, fcSttbfRMark);
        }

        public void SetLcbSttbfRMark(int lcbSttbfRMark)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.STTBFRMARK, lcbSttbfRMark);
        }

        /**
         * Return the offset to the PlcfHdd, in the table stream,
         * i.e. fcPlcfHdd
         */
        public int GetPlcfHddOffset()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFHDD);
        }
        /**
         * Return the size of the PlcfHdd, in the table stream,
         * i.e. lcbPlcfHdd
         */
        public int GetPlcfHddSize()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFHDD);
        }
        public void SetPlcfHddOffset(int fcPlcfHdd)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFHDD, fcPlcfHdd);
        }
        public void SetPlcfHddSize(int lcbPlcfHdd)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFHDD, lcbPlcfHdd);
        }

        public int GetFcSttbSavedBy()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.STTBSAVEDBY);
        }

        public int GetLcbSttbSavedBy()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.STTBSAVEDBY);
        }

        public void SetFcSttbSavedBy(int fcSttbSavedBy)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.STTBSAVEDBY, fcSttbSavedBy);
        }

        public void SetLcbSttbSavedBy(int fcSttbSavedBy)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.STTBSAVEDBY, fcSttbSavedBy);
        }

        public int GetModifiedLow()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLFLFO);
        }

        public int GetModifiedHigh()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLFLFO);
        }

        public void SetModifiedLow(int modifiedLow)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLFLFO, modifiedLow);
        }

        public void SetModifiedHigh(int modifiedHigh)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLFLFO, modifiedHigh);
        }


        /**
         * How many bytes of the main stream contain real data.
         */
        public int GetCbMac()
        {
            return _longHandler.GetLong(FIBLongHandler.CBMAC);
        }
        /**
         * Updates the count of the number of bytes in the
         * main stream which contain real data
         */
        public void SetCbMac(int cbMac)
        {
            _longHandler.SetLong(FIBLongHandler.CBMAC, cbMac);
        }
        /**
 * @return length of specified subdocument text stream in characters
 */
        public int GetSubdocumentTextStreamLength(SubdocumentType type)
        {
            return _longHandler.GetLong((int)type);
        }
        public void SetSubdocumentTextStreamLength(SubdocumentType type, int length)
        {
            if (length < 0)
                throw new ArgumentException(
                        "Subdocument length can't be less than 0 (passed value is "
                                + length + "). " + "If there is no subdocument "
                                + "length must be Set to zero.");

            _longHandler.SetLong((int)type, length);
        }
        /**
         * The count of CPs in the main document
         */
        public int GetCcpText()
        {
            return _longHandler.GetLong(FIBLongHandler.CCPTEXT);
        }
        /**
         * Updates the count of CPs in the main document
         */
        public void SetCcpText(int ccpText)
        {
            _longHandler.SetLong(FIBLongHandler.CCPTEXT, ccpText);
        }

        /**
         * The count of CPs in the footnote subdocument
         */
        public int GetCcpFtn()
        {
            return _longHandler.GetLong(FIBLongHandler.CCPFTN);
        }
        /**
         * Updates the count of CPs in the footnote subdocument
         */
        public void SetCcpFtn(int ccpFtn)
        {
            _longHandler.SetLong(FIBLongHandler.CCPFTN, ccpFtn);
        }

        /**
         * The count of CPs in the header story subdocument
         */
        public int GetCcpHdd()
        {
            return _longHandler.GetLong(FIBLongHandler.CCPHDD);
        }
        /**
         * Updates the count of CPs in the header story subdocument
         */
        public void SetCcpHdd(int ccpHdd)
        {
            _longHandler.SetLong(FIBLongHandler.CCPHDD, ccpHdd);
        }

        /**
         * The count of CPs in the comments (atn) subdocument
         */
        public int GetCcpAtn()
        {
            return _longHandler.GetLong(FIBLongHandler.CCPATN);
        }
        public int GetCcpCommentAtn()
        {
            return GetCcpAtn();
        }
        /**
         * Updates the count of CPs in the comments (atn) story subdocument
         */
        public void SetCcpAtn(int ccpAtn)
        {
            _longHandler.SetLong(FIBLongHandler.CCPATN, ccpAtn);
        }

        /**
         * The count of CPs in the end note subdocument
         */
        public int GetCcpEdn()
        {
            return _longHandler.GetLong(FIBLongHandler.CCPEDN);
        }
        /**
         * Updates the count of CPs in the end note subdocument
         */
        public void SetCcpEdn(int ccpEdn)
        {
            _longHandler.SetLong(FIBLongHandler.CCPEDN, ccpEdn);
        }

        /**
         * The count of CPs in the main document textboxes
         */
        public int GetCcpTxtBx()
        {
            return _longHandler.GetLong(FIBLongHandler.CCPTXBX);
        }
        /**
         * Updates the count of CPs in the main document textboxes
         */
        public void SetCcpTxtBx(int ccpTxtBx)
        {
            _longHandler.SetLong(FIBLongHandler.CCPTXBX, ccpTxtBx);
        }

        /**
         * The count of CPs in the header textboxes
         */
        public int GetCcpHdrTxtBx()
        {
            return _longHandler.GetLong(FIBLongHandler.CCPHDRTXBX);
        }
        /**
         * Updates the count of CPs in the header textboxes
         */
        public void SetCcpHdrTxtBx(int ccpTxtBx)
        {
            _longHandler.SetLong(FIBLongHandler.CCPHDRTXBX, ccpTxtBx);
        }


        public void ClearOffsetsSizes()
        {
            _fieldHandler.ClearFields();
        }

        public int GetFieldsPlcfOffset(FieldsDocumentPart part)
        {
            return _fieldHandler.GetFieldOffset((int)part);
        }

        public int GetFieldsPlcfLength(FieldsDocumentPart part)
        {
            return _fieldHandler.GetFieldSize((int)part);
        }

        public void SetFieldsPlcfOffset(FieldsDocumentPart part, int offSet)
        {
            _fieldHandler.SetFieldOffset((int)part, offSet);
        }

        public void SetFieldsPlcfLength(FieldsDocumentPart part, int length)
        {
            _fieldHandler.SetFieldSize((int)part, length);
        }

        [Obsolete]
        public int GetFcPlcffldAtn()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFFLDATN);
        }

        [Obsolete]
        public int GetLcbPlcffldAtn()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFFLDATN);
        }

        [Obsolete]
        public void SetFcPlcffldAtn(int offset)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFFLDATN, offset);
        }

        [Obsolete]
        public void SetLcbPlcffldAtn(int size)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFFLDATN, size);
        }

        [Obsolete]
        public int GetFcPlcffldEdn()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFFLDEDN);
        }

        [Obsolete]
        public int GetLcbPlcffldEdn()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFFLDEDN);
        }

        [Obsolete]
        public void SetFcPlcffldEdn(int offset)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFFLDEDN, offset);
        }

        [Obsolete]
        public void SetLcbPlcffldEdn(int size)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFFLDEDN, size);
        }

        [Obsolete]
        public int GetFcPlcffldFtn()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFFLDFTN);
        }

        [Obsolete]
        public int GetLcbPlcffldFtn()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFFLDFTN);
        }

        [Obsolete]
        public void SetFcPlcffldFtn(int offset)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFFLDFTN, offset);
        }

        [Obsolete]
        public void SetLcbPlcffldFtn(int size)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFFLDFTN, size);
        }

        [Obsolete]
        public int GetFcPlcffldHdr()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFFLDHDR);
        }

        [Obsolete]
        public int GetLcbPlcffldHdr()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFFLDHDR);
        }

        [Obsolete]
        public void SetFcPlcffldHdr(int offset)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFFLDHDR, offset);
        }

        [Obsolete]
        public void SetLcbPlcffldHdr(int size)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFFLDHDR, size);
        }

        [Obsolete]
        public int GetFcPlcffldHdrtxbx()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFFLDHDRTXBX);
        }

        [Obsolete]
        public int GetLcbPlcffldHdrtxbx()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFFLDHDRTXBX);
        }

        [Obsolete]
        public void SetFcPlcffldHdrtxbx(int offset)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFFLDHDRTXBX, offset);
        }

        [Obsolete]
        public void SetLcbPlcffldHdrtxbx(int size)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFFLDHDRTXBX, size);
        }

        [Obsolete]
        public int GetFcPlcffldMom()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFFLDMOM);
        }

        public int GetLcbPlcffldMom()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFFLDMOM);
        }

        [Obsolete]
        public void SetFcPlcffldMom(int offset)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFFLDMOM, offset);
        }

        [Obsolete]
        public void SetLcbPlcffldMom(int size)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFFLDMOM, size);
        }

        [Obsolete]
        public int GetFcPlcffldTxbx()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCFFLDTXBX);
        }

        [Obsolete]
        public int GetLcbPlcffldTxbx()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCFFLDTXBX);
        }

        [Obsolete]
        public void SetFcPlcffldTxbx(int offset)
        {
            _fieldHandler.SetFieldOffset(FIBFieldHandler.PLCFFLDTXBX, offset);
        }

        [Obsolete]
        public void SetLcbPlcffldTxbx(int size)
        {
            _fieldHandler.SetFieldSize(FIBFieldHandler.PLCFFLDTXBX, size);
        }


        public int GetFSPAPlcfOffset(FSPADocumentPart part)
        {
            return _fieldHandler.GetFieldOffset(part.GetFibFieldsField());
        }

        public int GetFSPAPlcfLength(FSPADocumentPart part)
        {
            return _fieldHandler.GetFieldSize(part.GetFibFieldsField());
        }

        public void SetFSPAPlcfOffset(FSPADocumentPart part, int offset)
        {
            _fieldHandler.SetFieldOffset(part.GetFibFieldsField(), offset);
        }

        public void SetFSPAPlcfLength(FSPADocumentPart part, int length)
        {
            _fieldHandler.SetFieldSize(part.GetFibFieldsField(), length);
        }

        [Obsolete]
        public int GetFcPlcspaMom()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.PLCSPAMOM);
        }

        [Obsolete]
        public int GetLcbPlcspaMom()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.PLCSPAMOM);
        }

        public int GetFcDggInfo()
        {
            return _fieldHandler.GetFieldOffset(FIBFieldHandler.DGGINFO);
        }

        public int GetLcbDggInfo()
        {
            return _fieldHandler.GetFieldSize(FIBFieldHandler.DGGINFO);
        }

        public int GetNotesDescriptorsOffset(NoteType noteType)
        {
            return _fieldHandler.GetFieldOffset(noteType
                    .GetFibDescriptorsFieldIndex());
        }

        public void SetNotesDescriptorsOffset(NoteType noteType, int offset)
        {
            _fieldHandler.SetFieldOffset(noteType.GetFibDescriptorsFieldIndex(),
                    offset);
        }

        public int GetNotesDescriptorsSize(NoteType noteType)
        {
            return _fieldHandler.GetFieldSize(noteType
                    .GetFibDescriptorsFieldIndex());
        }

        public void SetNotesDescriptorsSize(NoteType noteType, int offset)
        {
            _fieldHandler.SetFieldSize(noteType.GetFibDescriptorsFieldIndex(),
                    offset);
        }

        public int GetNotesTextPositionsOffset(NoteType noteType)
        {
            return _fieldHandler.GetFieldOffset(noteType
                    .GetFibTextPositionsFieldIndex());
        }

        public void SetNotesTextPositionsOffset(NoteType noteType, int offset)
        {
            _fieldHandler.SetFieldOffset(noteType.GetFibTextPositionsFieldIndex(),
                    offset);
        }

        public int GetNotesTextPositionsSize(NoteType noteType)
        {
            return _fieldHandler.GetFieldSize(noteType
                    .GetFibTextPositionsFieldIndex());
        }

        public void SetNotesTextPositionsSize(NoteType noteType, int offset)
        {
            _fieldHandler.SetFieldSize(noteType.GetFibTextPositionsFieldIndex(),
                    offset);
        }

        public void WriteTo(byte[] mainStream, HWPFStream tableStream)
        {
            //HWPFOutputStream mainDocument = sys.GetStream("WordDocument");
            //HWPFOutputStream tableStream = sys.GetStream("1Table");

            base.Serialize(mainStream, 0);

            int size = base.GetSize();
            _shortHandler.Serialize(mainStream);
            _longHandler.Serialize(mainStream, size + _shortHandler.SizeInBytes());
            _fieldHandler.WriteTo(mainStream,
              base.GetSize() + _shortHandler.SizeInBytes() + _longHandler.SizeInBytes(), tableStream);

        }

        public override int GetSize()
        {
            return base.GetSize() + _shortHandler.SizeInBytes() +
              _longHandler.SizeInBytes() + _fieldHandler.SizeInBytes();
        }
        //    public Object Clone()
        //    {
        //      try
        //      {
        //        return super.Clone();
        //      }
        //      catch (CloneNotSupportedException e)
        //      {
        //        e.printStackTrace();
        //        return null;
        //      }
        //    }
    }



}