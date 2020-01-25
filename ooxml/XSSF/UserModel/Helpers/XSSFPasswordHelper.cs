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

namespace NPOI.XSSF.UserModel.Helpers
{
    using System;
    using System.Globalization;
    using System.Xml;
    using System.Xml.XPath;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.POIFS.Crypt;
    using NPOI.Util;
    using Org.BouncyCastle.Security;

    public class XSSFPasswordHelper
    {
        private XSSFPasswordHelper()
        {
            // no instances of this static class
        }
        public static void SetPassword(CT_SheetProtection xobj, String password, HashAlgorithm hashAlgo, String prefix)
        {
            if (password == null)
            {
                xobj.password = null;
                xobj.algorithmName = null;
                xobj.hashValue = null;
                xobj.saltValue = null;
                xobj.spinCount = null;
                return;
            }
            if (hashAlgo == null)
            {
                int hash = CryptoFunctions.CreateXorVerifier1(password);
                xobj.password = String.Format("{0:X4}", hash).ToUpper();
            }
            else
            {
                SecureRandom random = new SecureRandom();
                byte[] salt = random.GenerateSeed(16);
                int spinCount = 100000;
                byte[] hash = CryptoFunctions.HashPassword(password, hashAlgo, salt, spinCount, false);
                xobj.algorithmName = hashAlgo.jceId;
                xobj.hashValue = Convert.ToBase64String(hash);
                xobj.saltValue = Convert.ToBase64String(salt);
                xobj.spinCount = "" + spinCount;
            }
        }
        /**
         * Sets the XORed or hashed password 
         *
         * @param xobj the xmlbeans object which Contains the password attributes
         * @param password the password, if null, the password attributes will be Removed
         * @param hashAlgo the hash algorithm, if null the password will be XORed
         * @param prefix the prefix of the password attributes, may be null
         */
        public static void SetPassword(XmlNode xobj, String password, HashAlgorithm hashAlgo, String prefix)
        {
            XPathNavigator cur = xobj.CreateNavigator();

            if (password == null)
            {
                //dose the prefix is namespace? check it!!!
                if (cur.MoveToAttribute("password", prefix)) cur.DeleteSelf();
                if (cur.MoveToAttribute("algorithmName", prefix)) cur.DeleteSelf();
                if (cur.MoveToAttribute("hashValue", prefix)) cur.DeleteSelf();
                if (cur.MoveToAttribute("saltValue", prefix)) cur.DeleteSelf();
                if (cur.MoveToAttribute("spinCount", prefix)) cur.DeleteSelf();
                return;
            }

            //cur.ToFirstContentToken();
            if (hashAlgo == null)
            {
                int hash = CryptoFunctions.CreateXorVerifier1(password);
                cur.CreateAttribute(prefix, "password", null, String.Format("{0:X4}", hash).ToUpper());
                //cur.InsertAttributeWithValue(GetAttrName(prefix, "password"),
                //                             String.Format("{0:X}", hash).ToUpper());
            }
            else
            {
                SecureRandom random = new SecureRandom();
                byte[] salt = random.GenerateSeed(16);

                // Iterations specifies the number of times the hashing function shall be iteratively run (using each
                // iteration's result as the input for the next iteration).
                int spinCount = 100000;

                // Implementation Notes List:
                // --> In this third stage, the reversed byte order legacy hash from the second stage shall
                //     be Converted to Unicode hex string representation
                byte[] hash = CryptoFunctions.HashPassword(password, hashAlgo, salt, spinCount, false);
                cur.CreateAttribute(prefix, "algorithmName", null, hashAlgo.jceId);
                cur.CreateAttribute(prefix, "hashValue", null, Convert.ToBase64String(hash));
                cur.CreateAttribute(prefix, "saltValue", null, Convert.ToBase64String(salt));
                cur.CreateAttribute(prefix, "spinCount", null, ""+spinCount);
                //cur.InsertAttributeWithValue(GetAttrName(prefix, "algorithmName"), hashAlgo.jceId);
                //cur.InsertAttributeWithValue(GetAttrName(prefix, "hashValue"), Convert.ToBase64String(hash));
                //cur.InsertAttributeWithValue(GetAttrName(prefix, "saltValue"), Convert.ToBase64String(salt));
                //cur.InsertAttributeWithValue(GetAttrName(prefix, "spinCount"), "" + spinCount);
            }
            //cur.Dispose();
        }

        public static bool ValidatePassword(CT_SheetProtection xobj, String password, String prefix)
        {
            if (password == null) return false;

            string xorHashVal = xobj.password;
            string algoName = xobj.algorithmName;
            string hashVal = xobj.hashValue;
            string saltVal = xobj.saltValue;
            string spinCount = xobj.spinCount;
            if (xorHashVal != null)
            {
                int hash1 = Int32.Parse(xorHashVal, NumberStyles.HexNumber);
                int hash2 = CryptoFunctions.CreateXorVerifier1(password);
                return hash1 == hash2;
            }
            else
            {
                if (hashVal == null || algoName == null || saltVal == null || spinCount == null)
                {
                    return false;
                }

                byte[] hash1 = Convert.FromBase64String(hashVal);
                HashAlgorithm hashAlgo = HashAlgorithm.FromString(algoName);
                byte[] salt = Convert.FromBase64String(saltVal);
                int spinCnt = Int32.Parse(spinCount);
                byte[] hash2 = CryptoFunctions.HashPassword(password, hashAlgo, salt, spinCnt, false);
                return Arrays.Equals(hash1, hash2);
            }
        }
        /**
         * Validates the password, i.e.
         * calculates the hash of the given password and Compares it against the stored hash
         *
         * @param xobj the xmlbeans object which Contains the password attributes
         * @param password the password, if null the method will always return false,
         *  even if there's no password Set
         * @param prefix the prefix of the password attributes, may be null
         * 
         * @return true, if the hashes match
         */
        public static bool ValidatePassword(XmlNode xobj, String password, String prefix)
        {
            // TODO: is "velvetSweatshop" the default password?
            if (password == null) return false;

            XPathNavigator cur = xobj.CreateNavigator();
            cur.MoveToAttribute("password", prefix);
            String xorHashVal = cur.Value;
            cur.MoveToAttribute("algorithmName", prefix);
            String algoName = cur.Value;
            cur.MoveToAttribute("hashValue", prefix);
            String hashVal = cur.Value;
            cur.MoveToAttribute("saltValue", prefix);
            String saltVal = cur.Value;
            cur.MoveToAttribute("spinCount", prefix);
            String spinCount = cur.Value;
            //cur.Dispose();

            if (xorHashVal != null)
            {
                int hash1 = Int32.Parse(xorHashVal, NumberStyles.HexNumber);
                int hash2 = CryptoFunctions.CreateXorVerifier1(password);
                return hash1 == hash2;
            }
            else
            {
                if (hashVal == null || algoName == null || saltVal == null || spinCount == null)
                {
                    return false;
                }

                byte[] hash1 = Convert.FromBase64String(hashVal);
                HashAlgorithm hashAlgo = HashAlgorithm.FromString(algoName);
                byte[] salt = Convert.FromBase64String(saltVal);
                int spinCnt = Int32.Parse(spinCount);
                byte[] hash2 = CryptoFunctions.HashPassword(password, hashAlgo, salt, spinCnt, false);
                return Arrays.Equals(hash1, hash2);
            }
        }


        private static XmlQualifiedName GetAttrName(String prefix, String name)
        {
            if (string.IsNullOrEmpty(prefix))
            {
                return new XmlQualifiedName(name);
            }
            else
            {
                return new XmlQualifiedName(prefix + char.ToUpper(name[0]) + name.Substring(1));
            }
        }
    }

}