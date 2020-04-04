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
using NPOI.POIFS.FileSystem;

namespace NPOI.POIFS.Crypt
{
    using NPOI.Util;
    using System;

    public interface IKey
    {
        string GetAlgorithm();

        string GetFormat();

        byte[] GetEncoded();
    }

    public interface ISecretKey: IKey
    {

    }
    public interface IPrivateKey: IKey
    {

    }
    public abstract class Encryptor
	{
		internal static string DEFAULT_POIFS_ENTRY = Decryptor.DEFAULT_POIFS_ENTRY;
		private ISecretKey secretKey;

		/**
	 * Return a output stream for encrypted data.
	 *
	 * @param dir the node to write to
	 * @return encrypted stream
	 */
		public abstract OutputStream GetDataStream(DirectoryNode dir);

		// for tests
		public abstract void ConfirmPassword(string password, byte[] keySpec, byte[] keySalt, byte[] verifier, byte[] verifierSalt , byte[] integritySalt);

		public abstract void ConfirmPassword(string password);

		public static Encryptor GetInstance(EncryptionInfo info)
		{
			return info.Encryptor;
		}

		public OutputStream GetDataStream(NPOIFSFileSystem fs)
		{
			return GetDataStream(fs.Root);
		}
        public OutputStream GetDataStream(OPOIFSFileSystem fs)
        {
            return GetDataStream(fs.Root);
        }

        public OutputStream GetDataStream(POIFSFileSystem fs)
		{
			return GetDataStream(fs.Root);
		}

		public ISecretKey GetSecretKey()
		{
			return secretKey;
		}

		protected void SetSecretKey(ISecretKey secretKey)
		{
			this.secretKey = secretKey;
		}
	}

}