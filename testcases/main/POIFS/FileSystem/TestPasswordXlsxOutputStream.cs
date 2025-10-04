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

/* ================================================================
 * About NPOI
 * Author: Tony Qu
 * Author's email: tonyqus (at) gmail.com
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors: Seijiro Ikehata (modeverv at gmail.com)
 *
 * ==============================================================*/

namespace TestCases.POIFS.FileSystem
{
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using System;
    using System.IO;
    using System.Linq;

    public class TestPasswordXlsxOutputStream
    {
        private static bool TestDecryption(string encryptedPath, string password)
        {
            try
            {
                byte[] decrypted = XlsxEncryptor.Decrypt(encryptedPath, password);


                if(decrypted.Length >= 2 && decrypted[0] == 0x50 && decrypted[1] == 0x4B)
                {
                    Console.WriteLine("  ✓ ok ZIP file（check PK signature）");
                }
                else
                {
                    Console.WriteLine("  ✗ no ZIP signature");
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine($"  ✗ error: {ex.Message}");
                return false;
            }

            return true;
        }

        private static bool CompareWithOriginal(string originalPath, string dotnetPath, string poiPath, string password)
        {
            try
            {
                byte[] original = File.ReadAllBytes(originalPath);
                byte[] dotnetDecrypted = XlsxEncryptor.Decrypt(dotnetPath, password);
                byte[] poiDecrypted = XlsxEncryptor.Decrypt(poiPath, password);

                Console.WriteLine($"original: {original.Length} bytes");
                Console.WriteLine($"dotnet: {dotnetDecrypted.Length} bytes");
                Console.WriteLine($"poi: {poiDecrypted.Length} bytes");

                bool dotnetMatch = CompareBytes(original, dotnetDecrypted);
                bool poiMatch = CompareBytes(original, poiDecrypted);

                Console.WriteLine($"\ndotnet and original: {(dotnetMatch ? "✓ same" : "✗ not same")}");
                Console.WriteLine($"poi and original: {(poiMatch ? "✓ same" : "✗ not same")}");
                if(!(dotnetMatch && poiMatch))
                {
                    return false;
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine($"compare error: {ex.Message}");
                return false;
            }

            return true;
        }

        private static bool CompareBytes(byte[] a, byte[] b)
        {
            if(a.Length != b.Length)
            {
                return false;
            }

            return !a.Where((t, i) => t != b[i]).Any();
        }

        [Test]
        public void TestEncryptor()
        {
            POIDataSamples samples = POIDataSamples.GetPOIFSInstance();
            string originalPath = samples.GetFileInfo("encrypt_original.xlsx").FullName;
            string poiPath = samples.GetFileInfo("encrypt_encrypted_by_poi_password_is_pass.xlsx").FullName;
            string projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory));
            string outputPath = Path.Combine(projectDir, "encrypt_encrypted_by_npoi.xlsx");
            string password = "pass";

            try
            {
                XlsxEncryptor.FromFileToFile(
                    originalPath,
                    outputPath,
                    password
                );
                // compare
                bool checkResult = TestDecryption(outputPath, password);
                ClassicAssert.IsTrue(checkResult);
                checkResult = TestDecryption(poiPath, password);
                ClassicAssert.IsTrue(checkResult);
                checkResult = CompareWithOriginal(originalPath, outputPath, poiPath, password);
                ClassicAssert.IsTrue(checkResult);
            }
            finally
            {
                try
                {
                    File.Delete(outputPath);
                }
                catch(Exception)
                {
                    // no-op
                }
            }
        }

        [Test]
        public void TestOutputStream()
        {
            string projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory));
            string outputPath = Path.Combine(projectDir, "protected_n.xlsx");
            string password = "pass";

            try
            {
                XSSFWorkbook wb = new();
                ISheet sheet = wb.CreateSheet("Sheet1");
                sheet.CreateRow(0).CreateCell(0).SetCellValue("Hello");
                using PasswordXlsxOutputStream outputStream = new(outputPath, password);
                wb.Write(outputStream);
                bool checkResult = TestDecryption(outputPath, password);
                ClassicAssert.IsTrue(checkResult);
            }
            finally
            {
                try
                {
                    File.Delete(outputPath);
                }
                catch(Exception)
                {
                    // no-op
                }
            }
        }
    }
}