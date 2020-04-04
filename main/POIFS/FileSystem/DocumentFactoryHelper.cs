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

namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// A small base class for the various factories, e.g. WorkbookFactory, SlideShowFactory to combine common code here.
    /// </summary>
    public class DocumentFactoryHelper
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
                    return new FilterInputStream1(d.GetDataStream(fs.Root), fs);
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

        private class FilterInputStream1 : FilterInputStream
        {
            NPOIFSFileSystem fs;
            public FilterInputStream1(InputStream input, NPOIFSFileSystem fs)
                : base(input)
            {
                this.fs = fs;
            }
            public override void Close()
            {
                fs.Close();
                base.Close();
            }
        }
        /// <summary>
        /// Checks that the supplied InputStream (which MUST support mark and reset, or be a PushbackInputStream) has a OOXML (zip) header at the start of it.
        /// If your InputStream does not support mark / reset, then wrap it in a PushBackInputStream, then be sure to always use that, and not the original!
        /// </summary>
        /// <param name="inp">An InputStream which supports either mark/reset, or is a PushbackInputStream</param>
        /// <returns></returns>
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

    }

}
