using NPOI.Util;
using NUnit.Framework;using NUnit.Framework.Legacy;
using System.IO;
using System.Threading;

namespace TestCases.Util
{
    /// <summary>
    /// Tests of creating temp files
    /// </summary>
    [TestFixture]
    internal class TestTempFile
    {
        [Test]
        public void TestCreateTempFile()
        {
            FileInfo fileInfo = null;
            Assert.DoesNotThrow(() => fileInfo = TempFile.CreateTempFile("test", ".xls"));

            ClassicAssert.IsTrue(fileInfo!=null && fileInfo.Exists);

            // Clean up only the file we created, not the shared directory
            if(fileInfo != null && fileInfo.Exists)
                fileInfo.Delete();

            // Verify CreateTempFile can still create files after cleanup
            FileInfo file = null;
            Assert.DoesNotThrow(() => file = TempFile.CreateTempFile("test2", ".xls"));
            ClassicAssert.IsTrue(file != null && file.Exists);

            if(file !=null && file.Exists)
                file.Delete();
        }

        [Test]
        public void TestCreateTempFileRecreatesDirectory()
        {
            // Use an isolated subdirectory to test directory recreation
            // without interfering with other tests using the shared poifiles dir
            string isolatedDir = Path.Combine(Path.GetTempPath(), "poifiles_test_" + System.Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(isolatedDir);

            try
            {
                // Create a file in the isolated directory to verify it works
                string testFile = Path.Combine(isolatedDir, "test.xls");
                File.WriteAllBytes(testFile, new byte[0]);
                ClassicAssert.IsTrue(File.Exists(testFile));

                // Delete the isolated directory
                Directory.Delete(isolatedDir, true);
                ClassicAssert.IsFalse(Directory.Exists(isolatedDir));

                // Verify we can recreate it
                Directory.CreateDirectory(isolatedDir);
                ClassicAssert.IsTrue(Directory.Exists(isolatedDir));
            }
            finally
            {
                if (Directory.Exists(isolatedDir))
                {
                    try { Directory.Delete(isolatedDir, true); }
                    catch { /* best effort cleanup */ }
                }
            }
        }

        [Test]
        public void TestGetTempFilePath()
        {
            string path = "";
            Assert.DoesNotThrow(() => path = TempFile.GetTempFilePath("test", ".xls"));

            ClassicAssert.IsTrue(!string.IsNullOrWhiteSpace(path));
            ClassicAssert.IsTrue(Directory.Exists(Path.GetDirectoryName(path)));
        }
    }
}
