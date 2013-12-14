using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LoadAndSaveBack
{
    class Program
    {
        enum RunMode
        { 
            Excel,
            Word
        }
        static void Main(string[] args)
        {
            if (args.Length != 2)
                return;

            string src = args[0];
            string target = args[1];
            

            RunMode mode= RunMode.Excel;
            if (src.Contains(".docx"))
                mode = RunMode.Word;

            if (mode == RunMode.Excel)
            {
                Stream rfs = File.OpenRead(src);
                IWorkbook workbook = new XSSFWorkbook(rfs);
                rfs.Close();
                using (FileStream fs = File.Create(target))
                {
                    workbook.Write(fs);
                }
            }
            else
            {
                Stream rfs = File.OpenRead(src);
                XWPFDocument workbook = new XWPFDocument(rfs);
                rfs.Close();
                using (FileStream fs = File.Create(target))
                {
                    workbook.Write(fs);
                }
            }
        }
    }
}
