/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using NPOI.POIFS.FileSystem;
using System.Text;


namespace NPOI.POIFS.Crypt
{
    public class EncryptionInfo
    {
        private int versionMajor;
        private int versionMinor;
        private int encryptionFlags;

        private EncryptionHeader header;
        private EncryptionVerifier verifier;

        public EncryptionInfo(POIFSFileSystem fs)
            : this(fs.Root)
        {
        }
        public EncryptionInfo(NPOIFSFileSystem fs)
            : this(fs.Root)
        {
        }
        public EncryptionInfo(DirectoryNode dir)
        {
            DocumentInputStream dis = dir.CreateDocumentInputStream("EncryptionInfo");
            versionMajor = dis.ReadShort();
            versionMinor = dis.ReadShort();

            encryptionFlags = dis.ReadInt();

            if (versionMajor == 4 && versionMinor == 4 && encryptionFlags == 0x40)
            {
                StringBuilder builder = new StringBuilder();
                byte[] xmlDescriptor = new byte[dis.Available()];
                dis.Read(xmlDescriptor);
                foreach (byte b in xmlDescriptor)
                    builder.Append((char)b);
                string descriptor = builder.ToString();
                header = new EncryptionHeader(descriptor);
                verifier = new EncryptionVerifier(descriptor);
            }
            else
            {
                int hSize = dis.ReadInt();
                header = new EncryptionHeader(dis);
                if (header.Algorithm == EncryptionHeader.ALGORITHM_RC4)
                {
                    verifier = new EncryptionVerifier(dis, 20);
                }
                else
                {
                    verifier = new EncryptionVerifier(dis, 32);
                }
            }
        }

        public int VersionMajor
        {
            get
            {
                return versionMajor;
            }
        }

        public int VersionMinor
        {
            get
            {
                return versionMinor;
            }
        }

        public int EncryptionFlags
        {
            get
            {
                return encryptionFlags;
            }
        }

        public EncryptionHeader Header
        {
            get
            {
                return header;
            }
        }

        public EncryptionVerifier Verifier
        {
            get
            {
                return verifier;
            }
        }

    }
}
