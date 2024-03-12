
namespace NPOI.Util
{
    using System;
    using System.IO;
    using System.Threading;

    public class TempFile
    {

        private static string dir;
        /**
         * Creates a temporary file.  Files are collected into one directory and by default are
         * deleted on exit from the VM.  Files can be kept by defining the system property
         * <c>poi.keep.tmp.files</c>.
         * 
         * Dont forget to close all files or it might not be possible to delete them.
         */
        public static FileInfo CreateTempFile(String prefix, String suffix)
        {

            if (dir == null)
            {
                dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "poifiles")).FullName;
            }
            // Generate a unique new filename 
            string file = Path.Combine(dir, prefix + Guid.NewGuid().ToString() + suffix);
            while (File.Exists(file))
            {
                file = Path.Combine(dir, prefix + Guid.NewGuid().ToString() + suffix);
                Thread.Sleep(1);
            }
            FileStream newFile = new FileStream(file, FileMode.CreateNew, FileAccess.ReadWrite);
            newFile.Close();

            return new FileInfo(file);
        }

        public static string GetTempFilePath(String prefix, String suffix)
        {
            if (dir == null)
            {
                dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "poifiles")).FullName;
            }

            Random rnd = new Random(DateTime.Now.Millisecond);
            rnd.Next();
            Thread.Sleep(10);
            //return prefix + rnd.Next() + suffix;
            string path = Path.Combine(dir, prefix + rnd.Next() + suffix);
            while(File.Exists(path))
            {
                Thread.Sleep(10);
                path = Path.Combine(dir, prefix + rnd.Next() + suffix);
            }
            return path;
        }
    }
}
