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

using NPOI.HSSF.Converter;
using NPOI.HSSF.UserModel;
using System;
using System.IO;

namespace NPOI.Examples.ConvertExcelToHtml
{
    class Program
    {
        static void Main(string[] args)
        {
            HSSFWorkbook workbook;
            //the excel file to convert
            string fileName = "19599-1.xls";
            fileName = Path.Combine(Environment.CurrentDirectory, fileName);
            workbook = ExcelToHtmlUtils.LoadXls(fileName);

            ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter();

            //set output parameter
            excelToHtmlConverter.OutputColumnHeaders = false;
            excelToHtmlConverter.OutputHiddenColumns = true;
            excelToHtmlConverter.OutputHiddenRows = true;
            excelToHtmlConverter.OutputLeadingSpacesAsNonBreaking = false;
            excelToHtmlConverter.OutputRowNumbers = true;
            excelToHtmlConverter.UseDivsToSpan = true;

            //process the excel file
            excelToHtmlConverter.ProcessWorkbook(workbook);

            //output the html file
            excelToHtmlConverter.Document.Save(Path.ChangeExtension(fileName, "html"));
        }
    }
}
