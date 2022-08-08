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

using NPOI.POIFS.Common;
using NPOI.POIFS.Crypt;
using NPOI.Util;
using System;
using System.IO;
using System.Text;

namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// A small base class for the various factories, e.g. WorkbookFactory, SlideShowFactory to combine common code here.
    /// </summary>
    public static class DocumentFactoryHelper
    {
        /// <summary>
        /// Wrap the OLE2 data in the NPOIFSFileSystem into a decrypted stream by using the given password.
        /// </summary>
        /// <param name="fs">The OLE2 stream for the document</param>
        /// <param name="password">The password, null if the default password should be used</param>
        /// <returns>A stream for reading the decrypted data</returns>
        /// <exception cref="System.IO.IOException">If an error occurs while decrypting or if the password does not match</exception>
        public static InputStream GetDecryptedStream(NPOIFSFileSystem fs, String password)

        {
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor d = Decryptor.GetInstance(info);

            try
            {
                bool passwordCorrect = false;
                if (password != null && d.VerifyPassword(password))
                {
                    passwordCorrect = true;
                }
                if (!passwordCorrect && d.VerifyPassword(Decryptor.DEFAULT_PASSWORD))
                {
                    passwordCorrect = true;
                }
                if (passwordCorrect)
                {
                    return d.GetDataStream(fs.Root);
                }
                else
                {
                    if (password != null)
                        throw new EncryptedDocumentException("Password incorrect");
                    else
                        throw new EncryptedDocumentException("The supplied spreadsheet is protected, but no password was supplied");
                }
            }
            catch (Exception e)
            {
                throw new IOException("password does not match", e);
            }
        }

        /// <summary>
        /// Checks that the supplied InputStream (which MUST support mark and reset, or be a PushbackInputStream) has a OOXML (zip) header at the start of it.
        /// If your InputStream does not support mark / reset, then wrap it in a PushBackInputStream, then be sure to always use that, and not the original!
        /// </summary>
        /// <param name="inp">An InputStream which supports either mark/reset, or is a PushbackInputStream</param>
        public static bool HasOOXMLHeader(Stream inp)
        {
            // We want to peek at the first 4 bytes
            //inp.mark(4);

            byte[] header = new byte[4];
            int bytesRead = IOUtils.ReadFully(inp, header);

            // Wind back those 4 bytes
            if (inp is PushbackStream)
            {
                PushbackStream pin = (PushbackStream)inp;
                pin.Position = pin.Position - 4;
                //pin.unread(header, 0, bytesRead);
            }
            else
            {
                inp.Position = 0;
            }

            // Did it match the ooxml zip signature?
            return (
                bytesRead == 4 &&
                header[0] == POIFSConstants.OOXML_FILE_HEADER[0] &&
                header[1] == POIFSConstants.OOXML_FILE_HEADER[1] &&
                header[2] == POIFSConstants.OOXML_FILE_HEADER[2] &&
                header[3] == POIFSConstants.OOXML_FILE_HEADER[3]
            );
        }

        /// <summary>
        /// Detects if a given office document is protected by a password or not.
        /// Supported formats: Word, Excel and PowerPoint (both legacy and OpenXml).
        /// </summary>
        /// <param name="fileName">Path to an office document.</param>
        /// <returns>True if document is protected by a password, false otherwise.</returns>
        public static bool IsPasswordProtected(string fileName)
        {
            using (var stream = File.OpenRead(fileName))
                return IsPasswordProtected(stream);
        }

        /// <summary>
        /// Detects if a given office document is protected by a password or not.
        /// Supported formats: Word, Excel and PowerPoint (both legacy and OpenXml).
        /// </summary>
        /// <param name="stream">Office document stream.</param>
        /// <returns>True if document is protected by a password, false otherwise.</returns>
        public static bool IsPasswordProtected(Stream stream)
        {
            return GetPasswordProtected(stream) != OfficeProtectType.Other;
        }

        /// <summary>
        /// Detects if a given office document is protected by a password or not.
        /// Supported formats: Word, Excel and PowerPoint (both legacy and OpenXml).
        /// </summary>
        /// <param name="stream">Office document stream.</param>
        /// <returns>True if document is protected by a password, false otherwise.</returns>
        public static OfficeProtectType GetPasswordProtected(Stream stream)
        {
            // minimum file size for office file is 4k
            if (stream.Length < 4096)
                return OfficeProtectType.Other;

            // read file header
            stream.Seek(0, SeekOrigin.Begin);
            var compObjHeader = new byte[0x20];
            ReadFromStream(stream, compObjHeader);

            // check if we have plain zip file
            if (compObjHeader[0] == 0x50 && compObjHeader[1] == 0x4b && compObjHeader[2] == 0x03 && compObjHeader[4] == 0x04)
            {
                // this is a plain OpenXml document (not encrypted)
                return OfficeProtectType.Other;
            }

            // check compound object magic bytes
            if (compObjHeader[0] != 0xD0 || compObjHeader[1] != 0xCF)
            {
                // unknown document format
                return OfficeProtectType.Other;
            }

            int sectionSizePower = compObjHeader[0x1E];
            if (sectionSizePower < 8 || sectionSizePower > 16)
            {
                // invalid section size
                return OfficeProtectType.Other;
            }
            int sectionSize = 2 << (sectionSizePower - 1);

            const int defaultScanLength = 32768;
            long scanLength = Math.Min(defaultScanLength, stream.Length);

            // read header part for scan
            stream.Seek(0, SeekOrigin.Begin);
            var header = new byte[scanLength];
            ReadFromStream(stream, header);

            // check if we detected password protection

            var protectType = ScanForPassword(stream, header, sectionSize);
            if (protectType != OfficeProtectType.Other)
                return protectType;

            // if not, try to scan footer as well

            // read footer part for scan
            stream.Seek(-scanLength, SeekOrigin.End);
            var footer = new byte[scanLength];
            ReadFromStream(stream, footer);

            // finally return the result
            return ScanForPassword(stream, footer, sectionSize);
        }

        private static OfficeProtectType ScanForPassword(Stream stream, byte[] buffer, int sectionSize)
        {
            const string afterNamePadding = "\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0";

            try
            {
                string bufferString = Encoding.ASCII.GetString(buffer, 0, buffer.Length);

                // try to detect password protection used in new OpenXml documents
                // by searching for "EncryptedPackage" or "EncryptedSummary" streams
                const string encryptedPackageName = "E\0n\0c\0r\0y\0p\0t\0e\0d\0P\0a\0c\0k\0a\0g\0e" + afterNamePadding;
                const string encryptedSummaryName = "E\0n\0c\0r\0y\0p\0t\0e\0d\0S\0u\0m\0m\0a\0r\0y" + afterNamePadding;
                if (bufferString.Contains(encryptedPackageName) ||
                    bufferString.Contains(encryptedSummaryName))
                {
                    return OfficeProtectType.ProtectedOOXML;
                }

                // try to detect password protection for legacy Office documents
                const int coBaseOffset = 0x200;
                const int sectionIdOffset = 0x74;

                // check for Word header
                const string wordDocumentName = "W\0o\0r\0d\0D\0o\0c\0u\0m\0e\0n\0t" + afterNamePadding;
                int headerOffset = bufferString.IndexOf(wordDocumentName, StringComparison.InvariantCulture);
                int sectionId;
                if (headerOffset >= 0)
                {
                    sectionId = BitConverter.ToInt32(buffer, headerOffset + sectionIdOffset);
                    int sectionOffset = coBaseOffset + (sectionId * sectionSize);
                    const int fibScanSize = 0x10;

                    if (sectionOffset < 0 || sectionOffset + fibScanSize > stream.Length)
                        return OfficeProtectType.Other; // invalid document

                    var fibHeader = new byte[fibScanSize];
                    stream.Seek(sectionOffset, SeekOrigin.Begin);
                    ReadFromStream(stream, fibHeader);
                    short properties = BitConverter.ToInt16(fibHeader, 0x0A);
                    // check for fEncrypted FIB bit
                    const short fEncryptedBit = 0x0100;
                    if ((properties & fEncryptedBit) == fEncryptedBit)
                    {
                        return OfficeProtectType.ProtectedOffice;
                    }
                    else
                    {
                        return OfficeProtectType.Other;
                    }
                }

                // check for Excel header
                const string workbookName = "W\0o\0r\0k\0b\0o\0o\0k" + afterNamePadding;
                headerOffset = bufferString.IndexOf(workbookName, StringComparison.InvariantCulture);
                if (headerOffset >= 0)
                {
                    sectionId = BitConverter.ToInt32(buffer, headerOffset + sectionIdOffset);
                    int sectionOffset = coBaseOffset + (sectionId * sectionSize);
                    const int streamScanSize = 0x100;
                    if (sectionOffset < 0 || sectionOffset + streamScanSize > stream.Length)
                        return OfficeProtectType.Other; // invalid document
                    var workbookStream = new byte[streamScanSize];
                    stream.Seek(sectionOffset, SeekOrigin.Begin);
                    ReadFromStream(stream, workbookStream);
                    short record = BitConverter.ToInt16(workbookStream, 0);
                    short recordSize = BitConverter.ToInt16(workbookStream, sizeof(short));
                    const short bofMagic = 0x0809;
                    const short eofMagic = 0x000A;
                    const short filePassMagic = 0x002F;
                    if (record != bofMagic)
                        return OfficeProtectType.Other; // invalid BOF
                                      // scan for FILEPASS record until the end of the buffer
                    int offset = (sizeof(short) * 2) + recordSize;
                    int recordsLeft = 16; // simple infinite loop check just in case
                    do
                    {
                        record = BitConverter.ToInt16(workbookStream, offset);
                        if (record == filePassMagic)
                            return OfficeProtectType.ProtectedOffice;
                        recordSize = BitConverter.ToInt16(workbookStream, sizeof(short) + offset);
                        offset += (sizeof(short) * 2) + recordSize;
                        recordsLeft--;
                    } while (record != eofMagic && recordsLeft > 0);
                }
            }
            catch (Exception ex)
            {
                // BitConverter exceptions may be related to document format problems
                // so we just treat them as "password not detected" result
                if (ex is ArgumentException)
                    return OfficeProtectType.Other;
                // respect all the rest exceptions
                throw;
            }

            return OfficeProtectType.Other;
        }

        private static void ReadFromStream(Stream stream, byte[] buffer)
        {
            int bytesRead, count = buffer.Length;
            while (count > 0 && (bytesRead = stream.Read(buffer, 0, count)) > 0)
                count -= bytesRead;
            if (count > 0) throw new EndOfStreamException();
        }

        public enum OfficeProtectType
        {
            ProtectedOOXML,
            ProtectedOffice,
            Other
        }
    }
}
