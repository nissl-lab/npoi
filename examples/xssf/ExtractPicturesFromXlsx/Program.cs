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
using System.Collections.Generic;
using System.Text;

using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using System.Collections;
using System.Drawing;
using NPOI.XSSF.UserModel;

namespace ExtractPicturesFromXls
{
    class Program
    {
        static void Main(string[] args)
        {
            FileStream file = File.OpenRead(@"clothes.xlsx");
            IWorkbook workbook = new XSSFWorkbook(file);

            IList pictures = workbook.GetAllPictures();
            int i = 0;
            foreach (IPictureData pic in pictures)
            {
                string ext = pic.SuggestFileExtension();
                if (ext.Equals("jpeg"))
                {
                    Image jpg = Image.FromStream(new MemoryStream(pic.Data));
                    jpg.Save(string.Format("pic{0}.jpg",i++));
                }
                else if (ext.Equals("png"))
                {
                    Image png = Image.FromStream(new MemoryStream(pic.Data));
                    png.Save(string.Format("pic{0}.png", i++));
                }

            }

        }
    }
}
