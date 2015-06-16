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
using NUnit.Framework;
using NPOI.HSSF.Extractor;
using TestCases.HSSF;
using System.Text.RegularExpressions;
namespace NPOI.XSSF.Extractor
{
/**
 * Tests for {@link XSSFEventBasedExcelExtractor}
 */
public class TestXSSFEventBasedExcelExtractor  {
    protected XSSFEventBasedExcelExtractor GetExtractor(String sampleName) throws Exception {
        return new XSSFEventBasedExcelExtractor(XSSFTestDataSamples.
                openSamplePackage(sampleName));
    }

    /**
     * Get text out of the simple file
     */
    public void TestGetSimpleText()  {
        // a very simple file
       XSSFEventBasedExcelExtractor extractor = GetExtractor("sample.xlsx");
        extractor.GetText();
        
        String text = extractor.GetText();
        Assert.IsTrue(text.Length > 0);
        
        // Check sheet names
        Assert.IsTrue(text.startsWith("Sheet1"));
        Assert.IsTrue(text.EndsWith("Sheet3\n"));
        
        // Now without, will have text
        extractor.SetIncludeSheetNames(false);
        text = extractor.GetText();
        String CHUNK1 =
            "Lorem\t111\n" + 
            "ipsum\t222\n" + 
            "dolor\t333\n" + 
            "sit\t444\n" + 
            "amet\t555\n" + 
            "consectetuer\t666\n" + 
            "adipiscing\t777\n" + 
            "elit\t888\n" + 
            "Nunc\t999\n";
        String CHUNK2 =
            "The quick brown fox jumps over the lazy dog\n" +
            "hello, xssf	hello, xssf\n" +
            "hello, xssf	hello, xssf\n" +
            "hello, xssf	hello, xssf\n" +
            "hello, xssf	hello, xssf\n";
        Assert.AreEqual(
                CHUNK1 + 
                "at\t4995\n" + 
                CHUNK2
                , text);
        
        // Now Get formulas not their values
        extractor.SetFormulasNotResults(true);
        text = extractor.GetText();
        Assert.AreEqual(
                CHUNK1 +
                "at\tSUM(B1:B9)\n" + 
                CHUNK2, text);
        
        // With sheet names too
        extractor.SetIncludeSheetNames(true);
        text = extractor.GetText();
        Assert.AreEqual(
                "Sheet1\n" +
                CHUNK1 +
                "at\tSUM(B1:B9)\n" + 
                "rich Test\n" +
                CHUNK2 +
                "Sheet3\n"
                , text);
        
        extractor.Close();
    }
    
    public void TestGetComplexText()  {
        // A fairly complex file
       XSSFEventBasedExcelExtractor extractor = GetExtractor("AverageTaxRates.xlsx");
        extractor.GetText();
        
        String text = extractor.GetText();
        Assert.IsTrue(text.Length > 0);
        
        // Might not have all formatting it should do!
        Assert.IsTrue(text.startsWith(
                        "Avgtxfull\n" +
                        "(iii) AVERAGE TAX RATES ON ANNUAL"	
        ));
        
        extractor.Close();
    }
    
   public void TestInlineStrings(){
      XSSFEventBasedExcelExtractor extractor = GetExtractor("InlineStrings.xlsx");
      extractor.SetFormulasNotResults(true);
      String text = extractor.GetText();

      // Numbers
      Assert.IsTrue("Unable to find expected word in text\n" + text, text.Contains("43"));
      Assert.IsTrue("Unable to find expected word in text\n" + text, text.Contains("22"));
      
      // Strings
      Assert.IsTrue("Unable to find expected word in text\n" + text, text.Contains("ABCDE"));
      Assert.IsTrue("Unable to find expected word in text\n" + text, text.Contains("Long Text"));
      
      // Inline Strings
      Assert.IsTrue("Unable to find expected word in text\n" + text, text.Contains("1st Inline String"));
      Assert.IsTrue("Unable to find expected word in text\n" + text, text.Contains("And More"));
      
      // Formulas
      Assert.IsTrue("Unable to find expected word in text\n" + text, text.Contains("A2"));
      Assert.IsTrue("Unable to find expected word in text\n" + text, text.Contains("A5-A$2"));
        
      extractor.Close();
   }
   
    /**
     * Test that we return pretty much the same as
     *  ExcelExtractor does, when we're both passed
     *  the same file, just saved as xls and xlsx
     */
    public void TestComparedToOLE2() {
        // A fairly simple file - ooxml
       XSSFEventBasedExcelExtractor ooxmlExtractor = GetExtractor("SampleSS.xlsx");

        ExcelExtractor ole2Extractor =
            new ExcelExtractor(HSSFTestDataSamples.OpenSampleWorkbook("SampleSS.xls"));
        
        POITextExtractor[] extractors =
            new POITextExtractor[] { ooxmlExtractor, ole2Extractor };
        for (int i = 0; i < extractors.Length; i++) {
            POITextExtractor extractor = extractors[i];
            
            String text = extractor.GetText().ReplaceAll("[\r\t]", "");
            Assert.IsTrue(text.startsWith("First Sheet\nTest spreadsheet\n2nd row2nd row 2nd column\n"));
            Pattern pattern = Pattern.compile(".*13(\\.0+)?\\s+Sheet3.*", Pattern.DOTALL);
            Matcher m = pattern.matcher(text);
            Assert.IsTrue(m.matches());			
        }
        
        ole2Extractor.Close();
        ooxmlExtractor.Close();
    }
}

}