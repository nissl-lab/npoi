/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ColorfulMatrixTable
{
    class Program
    {
        static IWorkbook wb;
        static void Main(string[] args)
        {
            InitializeWorkbook(args);

            ISheet sheet1 = wb.CreateSheet("Sheet1");

            int x = 1;
            for (int i = 0; i < 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    ICell cell = row.CreateCell(j);
                    if (x % 2 == 0)
                    {
                        //fill background with blue
                        ICellStyle style1 = wb.CreateCellStyle();
                        style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index2;
                        style1.FillPattern = FillPattern.SolidForeground;
                        cell.CellStyle = style1;
                    }
                    else
                    {
                        //fill background with yellow
                        ICellStyle style1 = wb.CreateCellStyle();
                        style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index2;
                        style1.FillPattern = FillPattern.SolidForeground;
                        cell.CellStyle = style1;
                    }
                    x++;
                }
            }

            WriteToFile();
        }

        static void WriteToFile()
        {
            string filename = "test.xls";
            if (wb is XSSFWorkbook) filename += "x";
            //Write the stream data of workbook to the root directory

            FileStream file = new FileStream(filename, FileMode.Create);
            wb.Write(file);
            file.Close();
        }

        static void InitializeWorkbook(string[] args)
        {
            if (args.Length > 0 && args[0].Equals("-xls"))
                wb = new HSSFWorkbook();
            else
                wb = new XSSFWorkbook();
        }
    }
}
