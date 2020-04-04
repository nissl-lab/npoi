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
using NPOI.POIFS.EventFileSystem;
using NPOI.Util;

namespace NPOI.POIFS.Crypt
{
    using System;
    using System.Text;
    using NPOI.POIFS.Crypt.Standard;
    using NPOI.POIFS.FileSystem;


    public class DataSpaceMapUtils
    {
        public static void AddDefaultDataSpace(DirectoryEntry dir)
        {
            DataSpaceMapEntry dsme = new DataSpaceMapEntry(
                    new int[] { 0 }
                  , new String[] { Decryptor.DEFAULT_POIFS_ENTRY }
                  , "StrongEncryptionDataSpace"
              );
            DataSpaceMap dsm = new DataSpaceMap(new DataSpaceMapEntry[] { dsme });
            CreateEncryptionEntry(dir, "\u0006DataSpaces/DataSpaceMap", dsm);

            DataSpaceDefInition dsd = new DataSpaceDefInition(new String[] { "StrongEncryptionTransform" });
            CreateEncryptionEntry(dir, "\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace", dsd);

            TransformInfoHeader tih = new TransformInfoHeader(
                  1
                , "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}"
                , "Microsoft.Container.EncryptionTransform"
                , 1, 0, 1, 0, 1, 0
            );
            IRMDSTransformInfo irm = new IRMDSTransformInfo(tih, 0, null);
            CreateEncryptionEntry(dir, "\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary", irm);

            DataSpaceVersionInfo dsvi = new DataSpaceVersionInfo("Microsoft.Container.DataSpaces", 1, 0, 1, 0, 1, 0);
            CreateEncryptionEntry(dir, "\u0006DataSpaces/Version", dsvi);
        }

        public static DocumentEntry CreateEncryptionEntry(DirectoryEntry dir, String path, EncryptionRecord out1)
        {
            String[] parts = path.Split("/".ToCharArray());
            for (int i = 0; i < parts.Length - 1; i++)
            {
                dir = dir.HasEntry(parts[i])
                    ? (DirectoryEntry)dir.GetEntry(parts[i])
                    : dir.CreateDirectory(parts[i]);
            }

            byte[] buf = new byte[5000];
            LittleEndianByteArrayOutputStream bos = new LittleEndianByteArrayOutputStream(buf, 0);
            out1.Write(bos);

            String fileName = parts[parts.Length - 1];

            if (dir.HasEntry(fileName))
            {
                dir.GetEntry(fileName).Delete();
            }

            return dir.CreateDocument(fileName, bos.WriteIndex, new POIFSWriterListenerImpl(buf));
            
        }
        public class POIFSWriterListenerImpl : POIFSWriterListener
        {
            byte[] buf;
            public POIFSWriterListenerImpl(byte[] buf)
            {
                this.buf = buf;
            }
            public void ProcessPOIFSWriterEvent(POIFSWriterEvent event1)
            {
                try
                {
                    event1.Stream.Write(buf, 0, event1.Limit);
                }
                catch (IOException e)
                {
                    throw new EncryptedDocumentException(e);
                }
            }
        }
        public class DataSpaceMap : EncryptionRecord
        {
            DataSpaceMapEntry[] entries;

            public DataSpaceMap(DataSpaceMapEntry[] entries)
            {
                this.entries = entries;
            }

            public DataSpaceMap(ILittleEndianInput is1)
            {
                //@SuppressWarnings("unused")
                int length = is1.ReadInt();
                int entryCount = is1.ReadInt();
                entries = new DataSpaceMapEntry[entryCount];
                for (int i = 0; i < entryCount; i++)
                {
                    entries[i] = new DataSpaceMapEntry(is1);
                }
            }

            public void Write(LittleEndianByteArrayOutputStream os)
            {
                os.WriteInt(8);
                os.WriteInt(entries.Length);
                foreach (DataSpaceMapEntry dsme in entries)
                {
                    dsme.Write(os);
                }
            }
        }

        public class DataSpaceMapEntry : EncryptionRecord
        {
            int[] referenceComponentType;
            String[] referenceComponent;
            String dataSpaceName;

            public DataSpaceMapEntry(int[] referenceComponentType, String[] referenceComponent, String dataSpaceName)
            {
                this.referenceComponentType = referenceComponentType;
                this.referenceComponent = referenceComponent;
                this.dataSpaceName = dataSpaceName;
            }

            public DataSpaceMapEntry(ILittleEndianInput is1)
            {

                int length = is1.ReadInt();
                int referenceComponentCount = is1.ReadInt();
                referenceComponentType = new int[referenceComponentCount];
                referenceComponent = new String[referenceComponentCount];
                for (int i = 0; i < referenceComponentCount; i++)
                {
                    referenceComponentType[i] = is1.ReadInt();
                    referenceComponent[i] = ReadUnicodeLPP4(is1);
                }
                dataSpaceName = ReadUnicodeLPP4(is1);
            }

            public void Write(LittleEndianByteArrayOutputStream os)
            {
                int start = os.WriteIndex;
                ILittleEndianOutput sizeOut = os.CreateDelayedOutput(LittleEndianConsts.INT_SIZE);
                os.WriteInt(referenceComponent.Length);
                for (int i = 0; i < referenceComponent.Length; i++)
                {
                    os.WriteInt(referenceComponentType[i]);
                    WriteUnicodeLPP4(os, referenceComponent[i]);
                }
                WriteUnicodeLPP4(os, dataSpaceName);
                sizeOut.WriteInt(os.WriteIndex - start);
            }
        }

        public class DataSpaceDefInition : EncryptionRecord
        {
            String[] transformer;

            public DataSpaceDefInition(String[] transformer)
            {
                this.transformer = transformer;
            }

            public DataSpaceDefInition(ILittleEndianInput is1)
            {
                int headerLength = is1.ReadInt();
                int transformReferenceCount = is1.ReadInt();
                transformer = new String[transformReferenceCount];
                for (int i = 0; i < transformReferenceCount; i++)
                {
                    transformer[i] = ReadUnicodeLPP4(is1);
                }
            }

            public void Write(LittleEndianByteArrayOutputStream bos)
            {
                bos.WriteInt(8);
                bos.WriteInt(transformer.Length);
                foreach (String str in transformer)
                {
                    WriteUnicodeLPP4(bos, str);
                }
            }
        }

        public class IRMDSTransformInfo : EncryptionRecord
        {
            TransformInfoHeader transformInfoHeader;
            int extensibilityHeader;
            String xrMLLicense;

            public IRMDSTransformInfo(TransformInfoHeader transformInfoHeader, int extensibilityHeader, String xrMLLicense)
            {
                this.transformInfoHeader = transformInfoHeader;
                this.extensibilityHeader = extensibilityHeader;
                this.xrMLLicense = xrMLLicense;
            }

            public IRMDSTransformInfo(ILittleEndianInput is1)
            {
                transformInfoHeader = new TransformInfoHeader(is1);
                extensibilityHeader = is1.ReadInt();
                xrMLLicense = ReadUtf8LPP4(is1);
                // finish with 0x04 (int) ???
            }

            public void Write(LittleEndianByteArrayOutputStream bos)
            {
                transformInfoHeader.Write(bos);
                bos.WriteInt(extensibilityHeader);
                WriteUtf8LPP4(bos, xrMLLicense);
                bos.WriteInt(4); // where does this 4 come from???
            }
        }

        public class TransformInfoHeader : EncryptionRecord
        {
            int transformType;
            String transformerId;
            String transformerName;
            int readerVersionMajor = 1, readerVersionMinor = 0;
            int updaterVersionMajor = 1, updaterVersionMinor = 0;
            int writerVersionMajor = 1, writerVersionMinor = 0;

            public TransformInfoHeader(
                int transformType,
                String transformerId,
                String transformerName,
                int readerVersionMajor, int readerVersionMinor,
                int updaterVersionMajor, int updaterVersionMinor,
                int writerVersionMajor, int writerVersionMinor
            )
            {
                this.transformType = transformType;
                this.transformerId = transformerId;
                this.transformerName = transformerName;
                this.readerVersionMajor = readerVersionMajor;
                this.readerVersionMinor = readerVersionMinor;
                this.updaterVersionMajor = updaterVersionMajor;
                this.updaterVersionMinor = updaterVersionMinor;
                this.writerVersionMajor = writerVersionMajor;
                this.writerVersionMinor = writerVersionMinor;
            }

            public TransformInfoHeader(ILittleEndianInput is1)
            {

                int length = is1.ReadInt();
                transformType = is1.ReadInt();
                transformerId = ReadUnicodeLPP4(is1);
                transformerName = ReadUnicodeLPP4(is1);
                readerVersionMajor = is1.ReadShort();
                readerVersionMinor = is1.ReadShort();
                updaterVersionMajor = is1.ReadShort();
                updaterVersionMinor = is1.ReadShort();
                writerVersionMajor = is1.ReadShort();
                writerVersionMinor = is1.ReadShort();
            }

            public void Write(LittleEndianByteArrayOutputStream bos)
            {
                int start = bos.WriteIndex;
                ILittleEndianOutput sizeOut = bos.CreateDelayedOutput(LittleEndianConsts.INT_SIZE);
                bos.WriteInt(transformType);
                WriteUnicodeLPP4(bos, transformerId);
                sizeOut.WriteInt(bos.WriteIndex - start);
                WriteUnicodeLPP4(bos, transformerName);
                bos.WriteShort(readerVersionMajor);
                bos.WriteShort(readerVersionMinor);
                bos.WriteShort(updaterVersionMajor);
                bos.WriteShort(updaterVersionMinor);
                bos.WriteShort(writerVersionMajor);
                bos.WriteShort(writerVersionMinor);
            }
        }

        public class DataSpaceVersionInfo : EncryptionRecord
        {
            String featureIdentifier;
            int readerVersionMajor = 1, readerVersionMinor = 0;
            int updaterVersionMajor = 1, updaterVersionMinor = 0;
            int writerVersionMajor = 1, writerVersionMinor = 0;

            public DataSpaceVersionInfo(ILittleEndianInput is1)
            {
                featureIdentifier = ReadUnicodeLPP4(is1);
                readerVersionMajor = is1.ReadShort();
                readerVersionMinor = is1.ReadShort();
                updaterVersionMajor = is1.ReadShort();
                updaterVersionMinor = is1.ReadShort();
                writerVersionMajor = is1.ReadShort();
                writerVersionMinor = is1.ReadShort();
            }

            public DataSpaceVersionInfo(
                String featureIdentifier,
                int readerVersionMajor, int readerVersionMinor,
                int updaterVersionMajor, int updaterVersionMinor,
                int writerVersionMajor, int writerVersionMinor
            )
            {
                this.featureIdentifier = featureIdentifier;
                this.readerVersionMajor = readerVersionMajor;
                this.readerVersionMinor = readerVersionMinor;
                this.updaterVersionMajor = updaterVersionMajor;
                this.updaterVersionMinor = updaterVersionMinor;
                this.writerVersionMajor = writerVersionMajor;
                this.writerVersionMinor = writerVersionMinor;
            }

            public void Write(LittleEndianByteArrayOutputStream bos)
            {
                WriteUnicodeLPP4(bos, featureIdentifier);
                bos.WriteShort(readerVersionMajor);
                bos.WriteShort(readerVersionMinor);
                bos.WriteShort(updaterVersionMajor);
                bos.WriteShort(updaterVersionMinor);
                bos.WriteShort(writerVersionMajor);
                bos.WriteShort(writerVersionMinor);
            }
        }

        public static String ReadUnicodeLPP4(ILittleEndianInput is1)
        {
            int length = is1.ReadInt();
            if (length % 2 != 0)
            {
                throw new EncryptedDocumentException(
                    "UNICODE-LP-P4 structure is a multiple of 4 bytes. "
                    + "If PAdding is present, it MUST be exactly 2 bytes long");
            }

            String result = StringUtil.ReadUnicodeLE(is1, length / 2);
            if (length % 4 == 2)
            {
                // PAdding (variable): A Set of bytes that MUST be of the correct size such that the size of the 
                // UNICODE-LP-P4 structure is a multiple of 4 bytes. If PAdding is present, it MUST be exactly 
                // 2 bytes long, and each byte MUST be 0x00.            
                is1.ReadShort();
            }

            return result;
        }

        public static void WriteUnicodeLPP4(ILittleEndianOutput os, String string1)
        {
            byte[] buf = StringUtil.GetToUnicodeLE(string1);
            os.WriteInt(buf.Length);
            os.Write(buf);
            if (buf.Length % 4 == 2)
            {
                os.WriteShort(0);
            }
        }

        public static String ReadUtf8LPP4(ILittleEndianInput is1)
        {
            int length = is1.ReadInt();
            if (length == 0 || length == 4)
            {
                //@SuppressWarnings("unused")
                int skip = is1.ReadInt(); // ignore
                return length == 0 ? null : "";
            }

            byte[] data = new byte[length];
            is1.ReadFully(data);

            // Padding (variable): A set of bytes that MUST be of correct size such that the size of the UTF-8-LP-P4
            // structure is a multiple of 4 bytes. If PAdding is present, each byte MUST be 0x00. If 
            // the length is exactly 0x00000000, this specifies a null string, and the entire structure uses 
            // exactly 4 bytes. If the length is exactly 0x00000004, this specifies an empty string, and the 
            // entire structure also uses exactly 4 bytes
            int scratchedBytes = length % 4;
            if (scratchedBytes > 0)
            {
                for (int i = 0; i < (4 - scratchedBytes); i++)
                {
                    is1.ReadByte();
                }
            }
            
            return Encoding.UTF8.GetString(data, 0, data.Length);
            //return new String(data, 0, data.length, Charset.forName("UTF-8"));
        }

        public static void WriteUtf8LPP4(ILittleEndianOutput os, String str)
        {
            if (str == null || "".Equals(str))
            {
                os.WriteInt(str == null ? 0 : 4);
                os.WriteInt(0);
            }
            else
            {

                byte[] buf = Encoding.UTF8.GetBytes(str);// str.GetBytes(Charset.ForName("UTF-8"));
                os.WriteInt(buf.Length);
                os.Write(buf);
                int scratchBytes = buf.Length % 4;
                if (scratchBytes > 0)
                {
                    for (int i = 0; i < (4 - scratchBytes); i++)
                    {
                        os.WriteByte(0);
                    }
                }
            }
        }

    }
}