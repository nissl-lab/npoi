
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
                dir = Directory.CreateDirectory(Path.GetTempPath() + @"\poifiles").FullName;
            }
            // Generate a unique new filename 
            Random rnd = new Random(DateTime.Now.Millisecond);
            Thread.Sleep(1);
            string file= dir + "\\" + prefix + rnd.Next() + suffix;
            while (File.Exists(file))
            {
                file = dir + "\\" + prefix + rnd.Next() + suffix;
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
                dir = Directory.CreateDirectory(Path.GetTempPath() + @"\poifiles").FullName;
            }
            Random rnd = new Random(DateTime.Now.Millisecond);
            Thread.Sleep(10);
            //return prefix + rnd.Next() + suffix;
            return dir + "\\" + prefix + rnd.Next() + suffix;
        }
    }
}
