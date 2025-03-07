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
using System.IO;
using NPOI.HWPF.Model.IO;
namespace NPOI.HWPF.Model
{


    /**
     * Holds information about document notes (footnotes or ending notes)
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {doc} com)
     */

    public class NotesTables
    {
        private PlexOfCps descriptors = new PlexOfCps(
                FootnoteReferenceDescriptor.GetSize());

        private NoteType noteType;

        private PlexOfCps textPositions = new PlexOfCps(0);

        public NotesTables(NoteType noteType)
        {
            this.noteType = noteType;
            textPositions
                    .AddProperty(new GenericPropertyNode(0, 1, Array.Empty<byte>()));
        }

        public NotesTables(NoteType noteType, byte[] tableStream,
                FileInformationBlock fib)
        {
            this.noteType = noteType;
            Read(tableStream, fib);
        }

        public GenericPropertyNode GetDescriptor(int index)
        {
            return descriptors.GetProperty(index);
        }

        public int GetDescriptorsCount()
        {
            return descriptors.Length;
        }

        public GenericPropertyNode GetTextPosition(int index)
        {
            return textPositions.GetProperty(index);
        }

        private void Read(byte[] tableStream, FileInformationBlock fib)
        {
            int referencesStart = fib.GetNotesDescriptorsOffset(noteType);
            int referencesLength = fib.GetNotesDescriptorsSize(noteType);

            if (referencesStart != 0 && referencesLength != 0)
                this.descriptors = new PlexOfCps(tableStream, referencesStart,
                        referencesLength, FootnoteReferenceDescriptor.GetSize());

            int textPositionsStart = fib.GetNotesTextPositionsOffset(noteType);
            int textPositionsLength = fib.GetNotesTextPositionsSize(noteType);

            if (textPositionsStart != 0 && textPositionsLength != 0)
                this.textPositions = new PlexOfCps(tableStream,
                        textPositionsStart, textPositionsLength, 0);
        }

        public void WriteRef(FileInformationBlock fib, HWPFStream tableStream)
        {
            if (descriptors == null || descriptors.Length == 0)
            {
                fib.SetNotesDescriptorsOffset(noteType, tableStream.Offset);
                fib.SetNotesDescriptorsSize(noteType, 0);
                return;
            }

            int start = tableStream.Offset;
            byte[] data = descriptors.ToByteArray();
            tableStream.Write(data);
            int end =tableStream.Offset;

            fib.SetNotesDescriptorsOffset(noteType, start);
            fib.SetNotesDescriptorsSize(noteType, end - start);
        }

        public void WriteTxt(FileInformationBlock fib, HWPFStream tableStream)
        {
            if (textPositions == null || textPositions.Length == 0)
            {
                fib.SetNotesTextPositionsOffset(noteType, tableStream.Offset);
                fib.SetNotesTextPositionsSize(noteType, 0);
                return;
            }

            int start = tableStream.Offset;
            byte[] data = textPositions.ToByteArray();            
            tableStream.Write(data);
            int end = tableStream.Offset;

            fib.SetNotesTextPositionsOffset(noteType, start);
            fib.SetNotesTextPositionsSize(noteType, end - start);
        }
    }


}



