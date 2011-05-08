
namespace NPOI.Util
{
    using System;
    using System.Configuration;
    using System.IO;

    public class TempFile
    {
        static DirectoryInfo dir;
        static Random rnd = new Random();

        /**
         * Creates a temporary file.  Files are collected into one directory and by default are
         * deleted on exit from the VM.  Files can be kept by defining the system property
         * <c>poi.keep.tmp.files</c>.
         * 
         * Dont forget to close all files or it might not be possible to delete them.
         */
        public static FileStream CreateTempFile(String prefix, String suffix)
        {
            //if (dir == null)
            //{
            //    dir = Directory.CreateDirectory(Path.GetTempPath()+@"\poifiles");               
            //}

            FileStream newFile = File.Open(prefix + rnd.Next() + suffix,FileMode.OpenOrCreate);
            
            return newFile;
        }

        public static string GetTempFilePath(String prefix, String suffix)
        {
            //if (dir == null)
            //{
            //    dir = Directory.CreateDirectory(Path.GetTempPath() + @"\poifiles");
            //}

            return prefix + rnd.Next() + suffix;
            //return dir.Name + "\\" + prefix + rnd.Next() + suffix;
        }
    }
}
