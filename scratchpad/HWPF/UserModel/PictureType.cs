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

        public static PictureType BMP = new PictureType("image/bmp", "bmp", new byte[][] { new byte[] { (byte)'B', (byte)'M' } });

        public static PictureType EMF = new PictureType("image/x-emf", "emf", new byte[][] { new byte[] { 0x01, 0x00, 0x00, 0x00 } });

        public static PictureType GIF = new PictureType("image/gif", "gif", new byte[][] { new byte[] { (byte)'G', (byte)'I', (byte)'F' } });

        public static PictureType JPEG = new PictureType("image/jpeg", "jpg", new byte[][] { new byte[] { (byte)0xFF, (byte)0xD8 } });

        public static PictureType PICT = new PictureType("image/pict", ".pict", Array.Empty<byte>()[]);

        public static PictureType PNG = new PictureType("image/png", "png", new byte[][] { new byte[]{ (byte) 0x89, 0x50, 0x4E, 0x47,
            0x0D, 0x0A, 0x1A, 0x0A } });

        public static PictureType TIFF = new PictureType("image/tiff", "tiff", new byte[][] { new byte[]{ 0x49, 0x49, 0x2A, 0x00 },
            new byte[]{ 0x4D, 0x4D, 0x00, 0x2A } });

        public static PictureType UNKNOWN = new PictureType("image/unknown", "", new byte[][] { new byte[] { } });

        public static PictureType WMF = new PictureType("image/x-wmf", "wmf", new byte[][] {
            new byte[]{ (byte) 0xD7, (byte) 0xCD, (byte) 0xC6, (byte) 0x9A, 0x00, 0x00 },
            new byte[]{ 0x01, 0x00, 0x09, 0x00, 0x00, 0x03 } });

        public static PictureType[] Values = new PictureType[] 
            { 
                PictureType.BMP,PictureType.EMF,PictureType.GIF,
                PictureType.JPEG,PictureType.PNG,PictureType.TIFF,
                PictureType.WMF,PictureType.UNKNOWN
            };

        public static PictureType FindMatchingType(byte[] pictureContent)
        {
            foreach (PictureType pictureType in PictureType.Values)
                for (int i = 0; i < pictureType.Signatures.Length; i++)
                {
                    if (MatchSignature(pictureContent, pictureType.Signatures[i]))
                        return pictureType;
                }

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


