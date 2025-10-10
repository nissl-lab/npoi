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

using NPOI.OpenXml4Net.OPC;
using NPOI.POIFS.Crypt;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using OpenMcdf;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace TestCases.POIFS.FileSystem
{
    public class TestEncryptXlsx
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
        public void TestPasswordXssfWorkbook()
        {
            string projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory));
            string outputPath = Path.Combine(projectDir, "protected_n.xlsx");
            string outputPathNormal = Path.Combine(projectDir, "normal_n.xlsx");
            string password = "pass";

            try
            {
                XSSFWorkbook wb = new();
                ISheet sheet = wb.CreateSheet("Sheet1");
                sheet.CreateRow(0).CreateCell(0).SetCellValue("Hello");

                using(POIFSFileSystem fs = new())
                {
                    EncryptionInfo info = new(EncryptionMode.AgileXlsx);
                    Encryptor enc = info.Encryptor;
                    enc.ConfirmPassword(password);
                    using(OutputStream os = enc.GetDataStream(fs))
                    {
                        wb.Write(os);
                    }

                    using(FileStream fos = File.Create(outputPath))
                    {
                        fs.WriteFileSystem(fos);
                    }
                }

                using(FileStream fos = File.Create(outputPathNormal))
                {
                    wb.Write(fos);
                }

                bool checkResult = TestDecryption(outputPath, password);
                ClassicAssert.IsTrue(checkResult);

                // check a normal version
                using(FileStream fis = File.OpenRead(outputPathNormal))
                {
                    XSSFWorkbook wbReloaded = new(fis);
                    ClassicAssert.AreEqual(1, wbReloaded.NumberOfSheets);
                    ClassicAssert.AreEqual("Sheet1", wbReloaded.GetSheetAt(0).SheetName);
                    ClassicAssert.AreEqual("Hello",
                        wbReloaded.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);
                }
            }
            finally
            {
                try
                {
                    File.Delete(outputPath);
                    File.Delete(outputPathNormal);
                }
                catch(Exception)
                {
                    // no-op
                }
            }
        }

        [Test]
        public void TestPoiCompatibleApi()
        {
            POIDataSamples samples = POIDataSamples.GetPOIFSInstance();
            string originalPath = samples.GetFileInfo("encrypt_original.xlsx").FullName;
            string projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory));
            string outputPath = Path.Combine(projectDir, "encrypt_encrypted_by_npoi.xlsx");
            string password = "pass";

            try
            {
                using(POIFSFileSystem fs = new())
                {
                    EncryptionInfo info = new(EncryptionMode.AgileXlsx);
                    Encryptor enc = info.Encryptor;
                    enc.ConfirmPassword(password);

                    using(OPCPackage opc = OPCPackage.Open(originalPath, PackageAccess.READ))
                    using(OutputStream os = enc.GetDataStream(fs))
                    {
                        opc.Save(os);
                    }

                    using(FileStream fos = File.Create(outputPath))
                    {
                        fs.WriteFileSystem(fos);
                    }
                }

                // check
                byte[] decrypted = XlsxEncryptor.Decrypt(outputPath, password);
                byte[] original = File.ReadAllBytes(originalPath);

                CollectionAssert.AreEqual(original, decrypted);
            }
            finally
            {
                if(File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }
            }
        }

        [Test]
        public void TestByteExactCompatibilityWithApachePoi()
        {
            POIDataSamples samples = POIDataSamples.GetPOIFSInstance();
            string originalPath = samples.GetFileInfo("encrypt_original.xlsx").FullName;
            string poiPath = samples.GetFileInfo("encrypt_encrypted_by_poi_password_is_pass.xlsx").FullName;
            string projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory));
            string npoiPath = Path.Combine(projectDir, "encrypt_encrypted_by_npoi.xlsx");
            string password = "pass";

            try
            {
                using(POIFSFileSystem fs = new())
                {
                    EncryptionInfo info = new(EncryptionMode.AgileXlsx);
                    Encryptor enc = info.Encryptor;
                    enc.ConfirmPassword(password);

                    using(OPCPackage opc = OPCPackage.Open(originalPath, PackageAccess.READ))
                    using(OutputStream os = enc.GetDataStream(fs))
                    {
                        opc.Save(os);
                    }

                    using(FileStream fos = File.Create(npoiPath))
                    {
                        fs.WriteFileSystem(fos);
                    }
                }

                byte[] decryptedFromPoi = XlsxEncryptor.Decrypt(poiPath, password);

                byte[] decryptedFromNpoi = XlsxEncryptor.Decrypt(npoiPath, password);

                CollectionAssert.AreEqual(decryptedFromPoi, decryptedFromNpoi);

                byte[] original = File.ReadAllBytes(originalPath);
                CollectionAssert.AreEqual(original, decryptedFromPoi);
                CollectionAssert.AreEqual(original, decryptedFromNpoi);
            }
            finally
            {
                if(File.Exists(npoiPath))
                {
                    File.Delete(npoiPath);
                }
            }
        }

        [Test]
        public void TestEncryptionStructureCompatibility()
        {
            POIDataSamples samples = POIDataSamples.GetPOIFSInstance();
            string originalPath = samples.GetFileInfo("encrypt_original.xlsx").FullName;
            string poiPath = samples.GetFileInfo("encrypt_encrypted_by_poi_password_is_pass.xlsx").FullName;
            string projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory));
            string npoiPath = Path.Combine(projectDir, "encrypt_encrypted_by_npoi.xlsx");
            string password = "pass";

            try
            {
                using(POIFSFileSystem fs = new())
                {
                    EncryptionInfo info = new(EncryptionMode.AgileXlsx);
                    Encryptor enc = info.Encryptor;
                    enc.ConfirmPassword(password);

                    using(OPCPackage opc = OPCPackage.Open(originalPath, PackageAccess.READ))
                    using(OutputStream os = enc.GetDataStream(fs))
                    {
                        opc.Save(os);
                    }

                    using(FileStream fos = File.Create(npoiPath))
                    {
                        fs.WriteFileSystem(fos);
                    }
                }

                using(RootStorage poiRoot = RootStorage.OpenRead(poiPath))
                using(RootStorage npoiRoot = RootStorage.OpenRead(npoiPath))
                {
                    string poiEncInfo = ReadEncryptionInfoXml(poiRoot);
                    string npoiEncInfo = ReadEncryptionInfoXml(npoiRoot);

                    ClassicAssert.IsNotNull(poiEncInfo);
                    ClassicAssert.IsNotNull(npoiEncInfo);
                }

                byte[] decryptedFromNpoi = XlsxEncryptor.Decrypt(npoiPath, password);
                byte[] original = File.ReadAllBytes(originalPath);
                CollectionAssert.AreEqual(original, decryptedFromNpoi);
            }
            finally
            {
                if(File.Exists(npoiPath))
                {
                    File.Delete(npoiPath);
                }
            }
        }

        private static string ReadEncryptionInfoXml(RootStorage root)
        {
            using CfbStream stream = root.OpenStream("EncryptionInfo");
            using BinaryReader reader = new(stream);
            reader.ReadUInt16(); // version major
            reader.ReadUInt16(); // version minor
            reader.ReadUInt32(); // flags
            byte[] xmlBytes = reader.ReadBytes((int) stream.Length - 8);
            return Encoding.UTF8.GetString(xmlBytes);
        }


        [Test]
        public void TestOutput()
        {
            string projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory));
            string outputPath = Path.Combine(projectDir, "protected_n.xlsx");
            string outputPathNormal = Path.Combine(projectDir, "normal_n.xlsx");
            string password = "pass";

            try
            {
                XSSFWorkbook wb = new();
                ISheet sheet = wb.CreateSheet("Sheet1");
                sheet.CreateRow(0).CreateCell(0).SetCellValue("Hello");

                using(FileStream fs = File.Create(outputPath))
                {
                    wb.Write(fs, password);
                }
                using(FileStream fs = File.Create(outputPathNormal))
                {
                    wb.Write(fs);
                }

                bool checkResult = TestDecryption(outputPath, password);
                ClassicAssert.IsTrue(checkResult);

                var decrypted = XlsxEncryptor.Decrypt(outputPath, password);
                        
                using var ms = new MemoryStream(decrypted);
                using var decryptedWb = new XSSFWorkbook(ms);
        
                var decryptedSheet = decryptedWb.GetSheetAt(0);
                var decryptedCell = decryptedSheet.GetRow(0).GetCell(0);
        
                ClassicAssert.AreEqual("Hello", decryptedCell.StringCellValue, "Cell value mismatch");
                ClassicAssert.AreEqual("Sheet1", decryptedSheet.SheetName, "Sheet name mismatch");
        
                Console.WriteLine("✓ Encryption and decryption test passed");
            }
            finally
            {
                try
                {
                    File.Delete(outputPath);
                    File.Delete(outputPathNormal);
                }
                catch(Exception)
                {
                    // no-op
                }
            }
        }
    }
}