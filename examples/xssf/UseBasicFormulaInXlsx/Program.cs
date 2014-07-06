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

using System;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;


namespace UseBasicFormulaInXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            IWorkbook workbook = new HSSFWorkbook();

            ISheet s1=workbook.CreateSheet("Sheet1");
            //set A2
            s1.CreateRow(1).CreateCell(0).SetCellValue(-5);
            //set B2
            s1.GetRow(1).CreateCell(1).SetCellValue(1111);
            //set C2
            s1.GetRow(1).CreateCell(2).SetCellValue(7.623);
            //set A3
            s1.CreateRow(2).CreateCell(0).SetCellValue(2.2);

            //set A4=A2+A3
            s1.CreateRow(3).CreateCell(0).CellFormula = "A2+A3";
            //set D2=SUM(A2:C2);
            s1.GetRow(1).CreateCell(3).CellFormula = "SUM(A2:C2)";
            //set A5=cos(5)+sin(10)
            s1.CreateRow(4).CreateCell(0).CellFormula="cos(5)+sin(10)";


            //create another sheet
            ISheet s2 = workbook.CreateSheet("Sheet2");
            //set cross-sheet reference
            s2.CreateRow(0).CreateCell(0).CellFormula = "Sheet1!A2+Sheet1!A3";
            IFormulaEvaluator e = WorkbookFactory.CreateFormulaEvaluator(workbook);
            var cell = e.Evaluate(cell);
            WriteToFile();
        }


        static HSSFWorkbook hssfworkbook;

        static void WriteToFile()
        {
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(@"test.xlsx", FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }

    }
}
