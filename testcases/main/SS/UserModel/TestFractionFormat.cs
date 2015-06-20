/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

using System;
using NPOI.SS.UserModel;
using TestCases.HSSF;
using NUnit.Framework;
using System.IO;
using System.Text.RegularExpressions;
namespace TestCases.SS.UserModel
{

    /**
     * Tests for the Fraction Formatting part of DataFormatter.
     * Largely taken from bug #54686
     */
    [TestFixture]
    public class TestFractionFormat
    {
        [Test]
        public void TestSingle()
        {
            FractionFormat f = new FractionFormat("", "##");
            string val = "321.321";
            String ret = f.Format(val);
            Assert.AreEqual("26027/81", ret);
        }
        [Test]
        public void TestTruthFile()
        {
            Stream truthFile = HSSFTestDataSamples.OpenSampleFileStream("54686_fraction_formats.txt");
            TextReader reader = new StreamReader(truthFile);
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("54686_fraction_formats.xls");
            ISheet sheet = wb.GetSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            // Skip over the header row
            String truthLine = reader.ReadLine();
            String[] headers = truthLine.Split("\t".ToCharArray());
            truthLine = reader.ReadLine();

            for (int i = 1; i < sheet.LastRowNum && truthLine != null; i++)
            {
                IRow r = sheet.GetRow(i);
                String[] truths = truthLine.Split("\t".ToCharArray());
                // Intentionally ignore the last column (tika-1132), for now
                for (short j = 3; j < 12; j++)
                {
                    ICell cell = r.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String truth = Clean(truths[j]);
                    String testKey = truths[0] + ":" + truths[1] + ":" + headers[j];
                    String formatted = Clean(formatter.FormatCellValue(cell, Evaluator));
                    if (truths.Length <= j)
                    {
                        continue;
                    }

                    
                    Assert.AreEqual(truth, formatted, testKey);
                }
                truthLine = reader.ReadLine();
            }
            reader.Close();
        }

        private String Clean(String s)
        {
            //s = s.Trim().Replace(" +", " ").Replace("- +", "-");
            s = Regex.Replace(s.Trim(), " +", " ");
            s = Regex.Replace(s, "- +", "-");
            return s;
        }
    }

}