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
using System;
using NPOI.Util;
using NPOI.HWPF.Model.IO;
namespace NPOI.HWPF.Model
{


    public class BookmarksTables
    {
        private PlexOfCps descriptorsFirst = new PlexOfCps(4);

        private PlexOfCps descriptorsLim = new PlexOfCps(0);

        private String[] names = Array.Empty<String>();

        public BookmarksTables(byte[] tableStream, FileInformationBlock fib)
        {
            Read(tableStream, fib);
        }

        public int GetBookmarksCount()
        {
            return descriptorsFirst.Length;
        }

        public GenericPropertyNode GetDescriptorFirst(int index)
        {
            return descriptorsFirst.GetProperty(index);
        }

        public int GetDescriptorFirstIndex(GenericPropertyNode descriptorFirst)
        {
            // TODO: very non-optimal
            return Arrays.AsList(descriptorsFirst.ToPropertiesArray()).IndexOf(
                    descriptorFirst);
        }

        public GenericPropertyNode GetDescriptorLim(int index)
        {
            return descriptorsLim.GetProperty(index);
        }

        public int GetDescriptorsFirstCount()
        {
            return descriptorsFirst.Length;
        }

        public int GetDescriptorsLimCount()
        {
            return descriptorsLim.Length;
        }

        public String GetName(int index)
        {
            return names[index];
        }

        public int GetNamesCount()
        {
            return names.Length;
        }

        private void Read(byte[] tableStream, FileInformationBlock fib)
        {
            int namesStart = fib.GetFcSttbfbkmk();
            int namesLength = fib.GetLcbSttbfbkmk();

            if (namesStart != 0 && namesLength != 0)
                this.names = SttbfUtils.Read(tableStream, namesStart);

            int firstDescriptorsStart = fib.GetFcPlcfbkf();
            int firstDescriptorsLength = fib.GetLcbPlcfbkf();
            if (firstDescriptorsStart != 0 && firstDescriptorsLength != 0)
                descriptorsFirst = new PlexOfCps(tableStream,
                        firstDescriptorsStart, firstDescriptorsLength,
                        BookmarkFirstDescriptor.GetSize());

            int limDescriptorsStart = fib.GetFcPlcfbkl();
            int limDescriptorsLength = fib.GetLcbPlcfbkl();
            if (limDescriptorsStart != 0 && limDescriptorsLength != 0)
                descriptorsLim = new PlexOfCps(tableStream, limDescriptorsStart,
                        limDescriptorsLength, 0);
        }

        public void SetName(int index, String name)
        {
            if (index < names.Length)
            {
                String[] newNames = new String[index + 1];
                Array.Copy(names, 0, newNames, 0, names.Length);
                names = newNames;
            }
            names[index] = name;
        }

        public void WritePlcfBkmkf(FileInformationBlock fib,
                HWPFStream tableStream)
        {
            if (descriptorsFirst == null || descriptorsFirst.Length == 0)
            {
                fib.SetFcPlcfbkf(0);
                fib.SetLcbPlcfbkf(0);
                return;
            }

            int start = tableStream.Offset;
            tableStream.Write(descriptorsFirst.ToByteArray());
            int end = tableStream.Offset;

            fib.SetFcPlcfbkf(start);
            fib.SetLcbPlcfbkf(end - start);
        }

        public void WritePlcfBkmkl(FileInformationBlock fib,
                HWPFStream tableStream)
        {
            if (descriptorsLim == null || descriptorsLim.Length == 0)
            {
                fib.SetFcPlcfbkl(0);
                fib.SetLcbPlcfbkl(0);
                return;
            }

            int start = tableStream.Offset;
            tableStream.Write(descriptorsLim.ToByteArray());
            int end = tableStream.Offset;

            fib.SetFcPlcfbkl(start);
            fib.SetLcbPlcfbkl(end - start);
        }

        public void WriteSttbfBkmk(FileInformationBlock fib,
                HWPFStream tableStream)
        {
            if (names == null || names.Length == 0)
            {
                fib.SetFcSttbfbkmk(0);
                fib.SetLcbSttbfbkmk(0);
                return;
            }

            int start = tableStream.Offset;
            SttbfUtils.Write(tableStream, names);
            int end = tableStream.Offset;

            fib.SetFcSttbfbkmk(start);
            fib.SetLcbSttbfbkmk(end - start);
        }
    }
}





