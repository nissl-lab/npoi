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
namespace NPOI.HWPF.UserModel
{

    /**
     * Picture types supported by MS Word format
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */
    public class PictureType
    {

        /*BMP( "image/bmp", "bmp", new byte[][] { { 'B', 'M' } } ),

    EMF( "image/x-emf", "emf", new byte[][] { { 0x01, 0x00, 0x00, 0x00 } } ),

    GIF( "image/gif", "gif", new byte[][] { { 'G', 'I', 'F' } } ),

    JPEG( "image/jpeg", "jpg", new byte[][] { { (byte) 0xFF, (byte) 0xD8 } } ),

    PNG( "image/png", "png", new byte[][] { { (byte) 0x89, 0x50, 0x4E, 0x47,
            0x0D, 0x0A, 0x1A, 0x0A } } ),

    TIFF( "image/tiff", "tiff", new byte[][] { { 0x49, 0x49, 0x2A, 0x00 },
            { 0x4D, 0x4D, 0x00, 0x2A } } ),

    UNKNOWN( "image/unknown", "", new byte[][] {} ),

    WMF( "image/x-wmf", "wmf", new byte[][] {
            { (byte) 0xD7, (byte) 0xCD, (byte) 0xC6, (byte) 0x9A, 0x00, 0x00 },
            { 0x01, 0x00, 0x09, 0x00, 0x00, 0x03 } } );
         * */
        public static PictureType FindMatchingType(byte[] pictureContent)
        {
            foreach (PictureType pictureType in PictureType.values())
                foreach (byte[] signature in pictureType.Signatures)
                    if (MatchSignature(pictureContent, signature))
                        return pictureType;

            // TODO: DIB, PICT
            return PictureType.UNKNOWN;
        }

        private static bool MatchSignature(byte[] pictureData, byte[] signature)
        {
            if (pictureData.Length < signature.Length)
                return false;

            for (int i = 0; i < signature.Length; i++)
                if (pictureData[i] != signature[i])
                    return false;

            return true;
        }

        private String _extension;

        private String _mime;

        private byte[][] _signatures;

        private PictureType(String mime, String extension, byte[][] signatures)
        {
            this._mime = mime;
            this._extension = extension;
            this._signatures = signatures;
        }

        public String Extension
        {
            get
            {
                return _extension;
            }
        }

        public String Mime
        {
            get
            {
                return _mime;
            }
        }

        public byte[][] Signatures
        {
            get
            {
                return _signatures;
            }
        }

        public bool MatchSignature(byte[] pictureData)
        {
            foreach (byte[] signature in Signatures)
                if (MatchSignature(signature, pictureData))
                    return true;
            return false;
        }
    }
}


