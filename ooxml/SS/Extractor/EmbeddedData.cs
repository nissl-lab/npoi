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

using System.Text.RegularExpressions;

namespace NPOI.SS.Extractor

{
    using NPOI.SS.UserModel;
    using System;

    /// <summary>
    /// A collection of embedded object informations and content
    /// </summary>
    public class EmbeddedData
    {
        private string _filename;
        private byte[] _embeddedData;
        private IShape _shape;
        private String _contentType = "binary/octet-stream";

        public EmbeddedData(String filename, byte[] embeddedData, String contentType)
        {
            Filename = filename;
            SetEmbeddedData(embeddedData);
            _contentType = contentType;
        }

        private Regex regFile = new Regex("[^/\\\\]*[/\\\\]", RegexOptions.Compiled);
        /// <summary>
        /// </summary>
        /// <return>filename </return>
        public string Filename
        {
            get { return _filename; }
            set
            {
                if (value == null)
                {
                    this._filename = "unknown.bin";
                }
                else
                {
                    this._filename = regFile.Replace(value, "").Trim();
                }
            }
        }

        /// <summary>
        /// </summary>
        /// <return>embedded object byte array </return>
        public byte[] GetEmbeddedData()
        {
            return _embeddedData;
        }

        /// <summary>
        /// Sets the embedded object as byte array
        /// </summary>
        /// <param name="embeddedData">the embedded object byte array</param>
        public void SetEmbeddedData(byte[] embeddedData)
        {
            this._embeddedData = (embeddedData == null) ? null : (byte[])embeddedData.Clone();
        }

        public IShape Shape
        {
            get { return _shape; }
            set { this._shape = value; }
        }
        /// <summary>
        /// content-/mime-type of the embedded object, the default (if unknown) is {@code binary/octet-stream}
        /// </summary>
        public string ContentType
        {
            get { return _contentType; }
            set { _contentType = value; }
        }
        

    }
}

