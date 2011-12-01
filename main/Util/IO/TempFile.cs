
namespace NPOI.Util
{
    using System;
    using System.Configuration;
    using System.IO;

    public class TempFile
    {
        

        /**
         * Creates a temporary file.  Files are collected into one directory and by default are
         * deleted on exit from the VM.  Files can be kept by defining the system property
         * <c>poi.keep.tmp.files</c>.
         * 
         * Dont forget to close all files or it might not be possible to delete them.
         */
        public static FileInfo CreateTempFile(String prefix, String suffix)
        {
            
            //if (dir == null)
            //{
            //    dir = Directory.CreateDirectory(Path.GetTempPath()+@"\poifiles");               
            //}
            Random rnd = new Random((int)DateTime.Now.Ticks);
            string file=prefix + rnd.Next() + suffix;
            FileStream newFile = File.Create(file);
            newFile.Close();

            return new FileInfo(file);
        }

        public static string GetTempFilePath(String prefix, String suffix)
        {
            //if (dir == null)
            //{
            //    dir = Directory.CreateDirectory(Path.GetTempPath() + @"\poifiles");
            //}
            Random rnd = new Random((int)DateTime.Now.Ticks);
            return prefix + rnd.Next() + suffix;
            //return dir.Name + "\\" + prefix + rnd.Next() + suffix;
        }
    }
}
