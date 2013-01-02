/*
 *  ====================================================================
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for Additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * ====================================================================
 */

using System.Collections.Generic;
using System;
using NPOI.HWPF.Model.IO;
namespace NPOI.HWPF.Model
{

    /**
     * This class provides access to all the fields Plex.
     * 
     * @author Cedric Bosdonnat <cbosdonnat@novell.com>
     * 
     */
    public class FieldsTables
    {
        // The size in bytes of the FLD data structure
        private static int FLD_SIZE = 2;

        private static List<PlexOfField> ToArrayList(PlexOfCps plexOfCps)
        {
            if (plexOfCps == null)
                return new List<PlexOfField>();

            List<PlexOfField> fields = new List<PlexOfField>(
                    plexOfCps.Length);
            for (int i = 0; i < plexOfCps.Length; i++)
            {
                GenericPropertyNode propNode = plexOfCps.GetProperty(i);
                PlexOfField plex = new PlexOfField(propNode);
                fields.Add(plex);
            }

            return fields;
        }

        private Dictionary<FieldsDocumentPart, PlexOfCps> _tables;

        public FieldsTables(byte[] tableStream, FileInformationBlock fib)
        {
            Array values = Enum.GetValues(typeof(FieldsDocumentPart));
            _tables = new Dictionary<FieldsDocumentPart, PlexOfCps>(
                    values.Length);

            foreach (FieldsDocumentPart part in values)
            {
                PlexOfCps plexOfCps = ReadPLCF(tableStream, fib, part);
                _tables.Add(part, plexOfCps);
            }
        }

        public List<PlexOfField> GetFieldsPLCF(FieldsDocumentPart part)
        {
            return ToArrayList(_tables[part]);
        }

        private PlexOfCps ReadPLCF(byte[] tableStream, FileInformationBlock fib,
                FieldsDocumentPart documentPart)
        {
            int start = fib.GetFieldsPlcfOffset(documentPart);
            int length = fib.GetFieldsPlcfLength(documentPart);

            if (start <= 0 || length <= 0)
                return null;

            return new PlexOfCps(tableStream, start, length, FLD_SIZE);
        }

        private int SavePlex(FileInformationBlock fib, FieldsDocumentPart part,
                PlexOfCps plexOfCps, HWPFStream outputStream)
        {
            if (plexOfCps == null || plexOfCps.Length == 0)
            {
                fib.SetFieldsPlcfOffset(part, outputStream.Offset);
                fib.SetFieldsPlcfLength(part, 0);
                return 0;
            }

            byte[] data = plexOfCps.ToByteArray();

            int start = outputStream.Offset;
            int length = data.Length;

            outputStream.Write(data);

            fib.SetFieldsPlcfOffset(part, start);
            fib.SetFieldsPlcfLength(part, length);

            return length;
        }

        public void Write(FileInformationBlock fib, HWPFStream tableStream)
        {
            Array values = Enum.GetValues(typeof(FieldsDocumentPart));
            foreach (FieldsDocumentPart part in values)
            {
                PlexOfCps plexOfCps = _tables[part];
                SavePlex(fib, part, plexOfCps, tableStream);
            }
        }

    }
}