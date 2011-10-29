/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.CF;
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Util;
    using NPOI.Util;
    using NPOI.SS.UserModel;

    /**
     * Tests the serialization and deserialization of the TestCFRuleRecord
     * class works correctly.
     *
     * @author Dmitriy Kumshayev 
     */
    [TestClass]
    public class TestCFRuleRecord
    {
        [TestMethod]
        public void TestCreateCFRuleRecord()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            CFRuleRecord record = CFRuleRecord.Create(workbook, "7");
            TestCFRuleRecord1(record);

            // Serialize
            byte[] SerializedRecord = record.Serialize();

            // Strip header
            byte[] recordData = new byte[SerializedRecord.Length - 4];
            Array.Copy(SerializedRecord, 4, recordData, 0, recordData.Length);

            // DeSerialize
            record = new CFRuleRecord(TestcaseRecordInputStream.Create(CFRuleRecord.sid, recordData));

            // Serialize again
            byte[] output = record.Serialize();

            // Compare
            Assert.AreEqual(recordData.Length + 4, output.Length, "Output size"); //includes sid+recordlength

            for (int i = 0; i < recordData.Length; i++)
            {
                Assert.AreEqual(recordData[i], output[i + 4], "CFRuleRecord doesn't match");
            }
        }

        private void TestCFRuleRecord1(CFRuleRecord record)
	    {
		    FontFormatting fontFormatting = new FontFormatting();
		    TestFontFormattingAccessors(fontFormatting);
		    Assert.IsFalse(record.ContainsFontFormattingBlock);
		    record.FontFormatting = (fontFormatting);
		    Assert.IsTrue(record.ContainsFontFormattingBlock);

		    BorderFormatting borderFormatting = new BorderFormatting();
		    TestBorderFormattingAccessors(borderFormatting);
		    Assert.IsFalse(record.ContainsBorderFormattingBlock);
		    record.BorderFormatting = (borderFormatting);
		    Assert.IsTrue(record.ContainsBorderFormattingBlock);

		    Assert.IsFalse(record.IsLeftBorderModified);
		    record.IsLeftBorderModified = (true);
		    Assert.IsTrue(record.IsLeftBorderModified);

		    Assert.IsFalse(record.IsRightBorderModified);
		    record.IsRightBorderModified = (true);
		    Assert.IsTrue(record.IsRightBorderModified);

		    Assert.IsFalse(record.IsTopBorderModified);
		    record.IsTopBorderModified = (true);
		    Assert.IsTrue(record.IsTopBorderModified);

		    Assert.IsFalse(record.IsBottomBorderModified);
		    record.IsBottomBorderModified = (true);
		    Assert.IsTrue(record.IsBottomBorderModified);

		    Assert.IsFalse(record.IsTopLeftBottomRightBorderModified);
		    record.IsTopLeftBottomRightBorderModified = (true);
		    Assert.IsTrue(record.IsTopLeftBottomRightBorderModified);

		    Assert.IsFalse(record.IsBottomLeftTopRightBorderModified);
		    record.IsBottomLeftTopRightBorderModified = (true);
		    Assert.IsTrue(record.IsBottomLeftTopRightBorderModified);


		    PatternFormatting patternFormatting = new PatternFormatting();
		    TestPatternFormattingAccessors(patternFormatting);
		    Assert.IsFalse(record.ContainsPatternFormattingBlock);
		    record.PatternFormatting = (patternFormatting);
		    Assert.IsTrue(record.ContainsPatternFormattingBlock);

		    Assert.IsFalse(record.IsPatternBackgroundColorModified);
		    record.IsPatternBackgroundColorModified = (true);
		    Assert.IsTrue(record.IsPatternBackgroundColorModified);

		    Assert.IsFalse(record.IsPatternColorModified);
		    record.IsPatternColorModified = (true);
		    Assert.IsTrue(record.IsPatternColorModified);

		    Assert.IsFalse(record.IsPatternStyleModified);
		    record.IsPatternStyleModified = (true);
		    Assert.IsTrue(record.IsPatternStyleModified);
	    }

        private void TestPatternFormattingAccessors(PatternFormatting patternFormatting)
        {
            patternFormatting.FillBackgroundColor = (HSSFColor.GREEN.index);
            Assert.AreEqual(HSSFColor.GREEN.index, patternFormatting.FillBackgroundColor);

            patternFormatting.FillForegroundColor = (HSSFColor.INDIGO.index);
            Assert.AreEqual(HSSFColor.INDIGO.index, patternFormatting.FillForegroundColor);

            patternFormatting.FillPattern = (PatternFormatting.DIAMONDS);
            Assert.AreEqual(PatternFormatting.DIAMONDS, patternFormatting.FillPattern);
        }

        private void TestBorderFormattingAccessors(BorderFormatting borderFormatting)
        {
            borderFormatting.IsBackwardDiagonalOn = (false);
            Assert.IsFalse(borderFormatting.IsBackwardDiagonalOn);
            borderFormatting.IsBackwardDiagonalOn = (true);
            Assert.IsTrue(borderFormatting.IsBackwardDiagonalOn);

            borderFormatting.BorderBottom = (BorderFormatting.BORDER_DOTTED);
            Assert.AreEqual(BorderFormatting.BORDER_DOTTED, borderFormatting.BorderBottom);

            borderFormatting.BorderDiagonal = (BorderFormatting.BORDER_MEDIUM);
            Assert.AreEqual(BorderFormatting.BORDER_MEDIUM, borderFormatting.BorderDiagonal);

            borderFormatting.BorderLeft = (BorderFormatting.BORDER_MEDIUM_DASH_DOT_DOT);
            Assert.AreEqual(BorderFormatting.BORDER_MEDIUM_DASH_DOT_DOT, borderFormatting.BorderLeft);

            borderFormatting.BorderRight = (BorderFormatting.BORDER_MEDIUM_DASHED);
            Assert.AreEqual(BorderFormatting.BORDER_MEDIUM_DASHED, borderFormatting.BorderRight);

            borderFormatting.BorderTop = (BorderFormatting.BORDER_HAIR);
            Assert.AreEqual(BorderFormatting.BORDER_HAIR, borderFormatting.BorderTop);

            borderFormatting.BottomBorderColor = (HSSFColor.AQUA.index);
            Assert.AreEqual(HSSFColor.AQUA.index, borderFormatting.BottomBorderColor);

            borderFormatting.DiagonalBorderColor = (HSSFColor.RED.index);
            Assert.AreEqual(HSSFColor.RED.index, borderFormatting.DiagonalBorderColor);

            Assert.IsFalse(borderFormatting.IsForwardDiagonalOn);
            borderFormatting.IsForwardDiagonalOn = (true);
            Assert.IsTrue(borderFormatting.IsForwardDiagonalOn);

            borderFormatting.LeftBorderColor = (HSSFColor.BLACK.index);
            Assert.AreEqual(HSSFColor.BLACK.index, borderFormatting.LeftBorderColor);

            borderFormatting.RightBorderColor = (HSSFColor.BLUE.index);
            Assert.AreEqual(HSSFColor.BLUE.index, borderFormatting.RightBorderColor);

            borderFormatting.TopBorderColor = (HSSFColor.GOLD.index);
            Assert.AreEqual(HSSFColor.GOLD.index, borderFormatting.TopBorderColor);
        }


        private void TestFontFormattingAccessors(FontFormatting fontFormatting)
        {
            // Check for defaults
            Assert.IsFalse(fontFormatting.IsEscapementTypeModified);
            Assert.IsFalse(fontFormatting.IsFontCancellationModified);
            Assert.IsFalse(fontFormatting.IsFontOutlineModified);
            Assert.IsFalse(fontFormatting.IsFontShadowModified);
            Assert.IsFalse(fontFormatting.IsFontStyleModified);
            Assert.IsFalse(fontFormatting.IsUnderlineTypeModified);
            Assert.IsFalse(fontFormatting.IsFontWeightModified);

            Assert.IsFalse(fontFormatting.IsBold);
            Assert.IsFalse(fontFormatting.IsItalic);
            Assert.IsFalse(fontFormatting.IsOutlineOn);
            Assert.IsFalse(fontFormatting.IsShadowOn);
            Assert.IsFalse(fontFormatting.IsStruckout);

            Assert.AreEqual(0, fontFormatting.EscapementType);
            Assert.AreEqual(-1, fontFormatting.FontColorIndex);
            Assert.AreEqual(-1, fontFormatting.FontHeight);
            Assert.AreEqual(0, fontFormatting.FontWeight);
            Assert.AreEqual(0, fontFormatting.UnderlineType);

            fontFormatting.IsBold = (true);
            Assert.IsTrue(fontFormatting.IsBold);
            fontFormatting.IsBold = (false);
            Assert.IsFalse(fontFormatting.IsBold);

            fontFormatting.EscapementType = FontSuperScript.SUB;
            Assert.AreEqual(FontFormatting.SS_SUB, fontFormatting.EscapementType);
            fontFormatting.EscapementType = FontSuperScript.SUPER;
            Assert.AreEqual(FontFormatting.SS_SUPER, fontFormatting.EscapementType);
            fontFormatting.EscapementType = (FontFormatting.SS_NONE);
            Assert.AreEqual(FontFormatting.SS_NONE, fontFormatting.EscapementType);

            fontFormatting.IsEscapementTypeModified = (false);
            Assert.IsFalse(fontFormatting.IsEscapementTypeModified);
            fontFormatting.IsEscapementTypeModified = (true);
            Assert.IsTrue(fontFormatting.IsEscapementTypeModified);

            fontFormatting.IsFontWeightModified = (false);
            Assert.IsFalse(fontFormatting.IsFontWeightModified);
            fontFormatting.IsFontWeightModified = (true);
            Assert.IsTrue(fontFormatting.IsFontWeightModified);

            fontFormatting.IsFontCancellationModified = (false);
            Assert.IsFalse(fontFormatting.IsFontCancellationModified);
            fontFormatting.IsFontCancellationModified = (true);
            Assert.IsTrue(fontFormatting.IsFontCancellationModified);

            fontFormatting.FontColorIndex = ((short)10);
            Assert.AreEqual(10, fontFormatting.FontColorIndex);

            fontFormatting.FontHeight = (100);
            Assert.AreEqual(100, fontFormatting.FontHeight);

            fontFormatting.IsFontOutlineModified = (false);
            Assert.IsFalse(fontFormatting.IsFontOutlineModified);
            fontFormatting.IsFontOutlineModified = (true);
            Assert.IsTrue(fontFormatting.IsFontOutlineModified);

            fontFormatting.IsFontShadowModified = (false);
            Assert.IsFalse(fontFormatting.IsFontShadowModified);
            fontFormatting.IsFontShadowModified = (true);
            Assert.IsTrue(fontFormatting.IsFontShadowModified);

            fontFormatting.IsFontStyleModified = (false);
            Assert.IsFalse(fontFormatting.IsFontStyleModified);
            fontFormatting.IsFontStyleModified = (true);
            Assert.IsTrue(fontFormatting.IsFontStyleModified);

            fontFormatting.IsItalic = (false);
            Assert.IsFalse(fontFormatting.IsItalic);
            fontFormatting.IsItalic = (true);
            Assert.IsTrue(fontFormatting.IsItalic);

            fontFormatting.IsOutlineOn = (false);
            Assert.IsFalse(fontFormatting.IsOutlineOn);
            fontFormatting.IsOutlineOn = (true);
            Assert.IsTrue(fontFormatting.IsOutlineOn);

            fontFormatting.IsShadowOn = (false);
            Assert.IsFalse(fontFormatting.IsShadowOn);
            fontFormatting.IsShadowOn = (true);
            Assert.IsTrue(fontFormatting.IsShadowOn);

            fontFormatting.IsStruckout = (false);
            Assert.IsFalse(fontFormatting.IsStruckout);
            fontFormatting.IsStruckout = (true);
            Assert.IsTrue(fontFormatting.IsStruckout);

            fontFormatting.UnderlineType = FontUnderlineType.DOUBLE_ACCOUNTING;
            Assert.AreEqual(FontFormatting.U_DOUBLE_ACCOUNTING, fontFormatting.UnderlineType);

            fontFormatting.IsUnderlineTypeModified = (false);
            Assert.IsFalse(fontFormatting.IsUnderlineTypeModified);
            fontFormatting.IsUnderlineTypeModified = (true);
            Assert.IsTrue(fontFormatting.IsUnderlineTypeModified);
        }
        [TestMethod]
        public void TestWrite() {
		    HSSFWorkbook workbook = new HSSFWorkbook();
		    CFRuleRecord rr = CFRuleRecord.Create(workbook, ComparisonOperator.BETWEEN, "5", "10");

		    PatternFormatting patternFormatting = new PatternFormatting();
		    patternFormatting.FillPattern=(PatternFormatting.BRICKS);
		    rr.PatternFormatting=(patternFormatting);

		    byte[] data = rr.Serialize();
		    Assert.AreEqual(26, data.Length);
		    Assert.AreEqual(3, LittleEndian.GetShort(data, 6));
		    Assert.AreEqual(3, LittleEndian.GetShort(data, 8));

		    int flags = LittleEndian.GetInt(data, 10);
		    Assert.AreEqual(0x00380000, flags & 0x00380000,"unused flags should be 111");
		    Assert.AreEqual(0, flags & 0x03C00000,"undocumented flags should be 0000"); // Otherwise Excel s unhappy
		    // check all remaining flag bits (some are not well understood yet)
		    Assert.AreEqual(0x203FFFFF, flags);
	    }

        private static byte[] DATA_REFN = {
        // formula extracted from bugzilla 45234 att 22141
            1, 3, 
            9, // formula 1 length 
            0, 0, 0, unchecked((byte)-1), unchecked((byte)-1), 63, 32, 2, unchecked((byte)-128), 0, 0, 0, 5,
            // formula 1: "=B3=1" (formula is relative to B4)
            76, unchecked((byte)-1), unchecked((byte)-1), 0, unchecked((byte)-64), // tRefN(B1)
            30, 1, 0,	
            11,	
        };

        /**
         * tRefN and tAreaN tokens must be preserved when re-serializing conditional format formulas
         */
        [TestMethod]
        public void TestReserializeRefNTokens() 
        {
    		
		    RecordInputStream is1 = TestcaseRecordInputStream.Create (CFRuleRecord.sid, DATA_REFN);
		    CFRuleRecord rr = new CFRuleRecord(is1);
		    Ptg[] ptgs = rr.ParsedExpression1;
		    Assert.AreEqual(3, ptgs.Length);
		    if (ptgs[0] is RefPtg) {
			    throw new AssertFailedException("Identified bug 45234");
		    }
		    Assert.AreEqual(typeof(RefNPtg), ptgs[0].GetType());
		    RefNPtg refNPtg = (RefNPtg) ptgs[0];
		    Assert.IsTrue(refNPtg.IsColRelative);
		    Assert.IsTrue(refNPtg.IsRowRelative);
    		    
		    byte[] data = rr.Serialize();
    		
		    if (!CompareArrays(DATA_REFN, 0, data, 4, DATA_REFN.Length)) {
			    Assert.Fail("Did not re-serialize correctly");
		    }
	    }

        private static bool CompareArrays(byte[] arrayA, int offsetA, byte[] arrayB, int offsetB, int length)
        {

            if (offsetA + length > arrayA.Length)
            {
                return false;
            }
            if (offsetB + length > arrayB.Length)
            {
                return false;
            }
            for (int i = 0; i < length; i++)
            {
                if (arrayA[i + offsetA] != arrayB[i + offsetB])
                {
                    return false;
                }
            }
            return true;
        }

    }
}