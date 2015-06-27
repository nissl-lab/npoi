
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
            string file=prefix + rnd.Next() + suffix;
            if (File.Exists(file))
            {
                // That name is already taken, try another
                file = CreateTempFile(prefix, suffix).FullName;
            }
            FileStream newFile = File.Create(file);
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
            return prefix + rnd.Next() + suffix;
            //return dir.Name + "\\" + prefix + rnd.Next() + suffix;
        }
    }
}
