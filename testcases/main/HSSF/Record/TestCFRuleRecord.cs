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
    using NUnit.Framework;
    using NUnit.Framework.Constraints;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.CF;
    using NPOI.SS.Formula;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Util;
    using NPOI.Util;
    using NPOI.SS.UserModel;
    using NPOI.SS.Formula.PTG;

    /**
     * Tests the serialization and deserialization of the TestCFRuleRecord
     * class works correctly.
     *
     * @author Dmitriy Kumshayev 
     */
    [TestFixture]
    public class TestCFRuleRecord
    {
        [Test]
        public void TestConstructors()
        {
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();

            CFRuleRecord rule1 = CFRuleRecord.Create((HSSFSheet)sheet, "7");
            Assert.AreEqual(CFRuleRecord.CONDITION_TYPE_FORMULA, rule1.ConditionType);
            Assert.AreEqual((byte)ComparisonOperator.NoComparison, rule1.ComparisonOperation);
            Assert.IsNotNull(rule1.ParsedExpression1);
            Assert.AreSame(Ptg.EMPTY_PTG_ARRAY, rule1.ParsedExpression2);

            CFRuleRecord rule2 = CFRuleRecord.Create((HSSFSheet)sheet, (byte)ComparisonOperator.Between, "2", "5");
            Assert.AreEqual(CFRuleRecord.CONDITION_TYPE_CELL_VALUE_IS, rule2.ConditionType);
            Assert.AreEqual((byte)ComparisonOperator.Between, rule2.ComparisonOperation);
            Assert.IsNotNull(rule2.ParsedExpression1);
            Assert.IsNotNull(rule2.ParsedExpression2);

            CFRuleRecord rule3 = CFRuleRecord.Create((HSSFSheet)sheet, (byte)ComparisonOperator.Equal, null, null);
            Assert.AreEqual(CFRuleRecord.CONDITION_TYPE_CELL_VALUE_IS, rule3.ConditionType);
            Assert.AreEqual((byte)ComparisonOperator.Equal, rule3.ComparisonOperation);
            Assert.AreSame(Ptg.EMPTY_PTG_ARRAY, rule3.ParsedExpression2);
            Assert.AreSame(Ptg.EMPTY_PTG_ARRAY, rule3.ParsedExpression2);
        }
        [Test]
        public void TestCreateCFRuleRecord()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet();
            CFRuleRecord record = CFRuleRecord.Create(sheet, "7");
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
            patternFormatting.FillBackgroundColor = (HSSFColor.Green.Index);
            Assert.AreEqual(HSSFColor.Green.Index, patternFormatting.FillBackgroundColor);

            patternFormatting.FillForegroundColor = (HSSFColor.Indigo.Index);
            Assert.AreEqual(HSSFColor.Indigo.Index, patternFormatting.FillForegroundColor);

            patternFormatting.FillPattern = FillPattern.Diamonds;
            Assert.AreEqual(FillPattern.Diamonds, patternFormatting.FillPattern);
        }

        private void TestBorderFormattingAccessors(BorderFormatting borderFormatting)
        {
            borderFormatting.IsBackwardDiagonalOn = (false);
            Assert.IsFalse(borderFormatting.IsBackwardDiagonalOn);
            borderFormatting.IsBackwardDiagonalOn = (true);
            Assert.IsTrue(borderFormatting.IsBackwardDiagonalOn);

            borderFormatting.BorderBottom = BorderStyle.Dotted;
            Assert.AreEqual(BorderStyle.Dotted, borderFormatting.BorderBottom);

            borderFormatting.BorderDiagonal = (BorderStyle.Medium);
            Assert.AreEqual(BorderStyle.Medium, borderFormatting.BorderDiagonal);

            borderFormatting.BorderLeft = (BorderStyle.MediumDashDotDot);
            Assert.AreEqual(BorderStyle.MediumDashDotDot, borderFormatting.BorderLeft);

            borderFormatting.BorderRight = (BorderStyle.MediumDashed);
            Assert.AreEqual(BorderStyle.MediumDashed, borderFormatting.BorderRight);

            borderFormatting.BorderTop = (BorderStyle.Hair);
            Assert.AreEqual(BorderStyle.Hair, borderFormatting.BorderTop);

            borderFormatting.BottomBorderColor = (HSSFColor.Aqua.Index);
            Assert.AreEqual(HSSFColor.Aqua.Index, borderFormatting.BottomBorderColor);

            borderFormatting.DiagonalBorderColor = (HSSFColor.Red.Index);
            Assert.AreEqual(HSSFColor.Red.Index, borderFormatting.DiagonalBorderColor);

            Assert.IsFalse(borderFormatting.IsForwardDiagonalOn);
            borderFormatting.IsForwardDiagonalOn = (true);
            Assert.IsTrue(borderFormatting.IsForwardDiagonalOn);

            borderFormatting.LeftBorderColor = (HSSFColor.Black.Index);
            Assert.AreEqual(HSSFColor.Black.Index, borderFormatting.LeftBorderColor);

            borderFormatting.RightBorderColor = (HSSFColor.Blue.Index);
            Assert.AreEqual(HSSFColor.Blue.Index, borderFormatting.RightBorderColor);

            borderFormatting.TopBorderColor = (HSSFColor.Gold.Index);
            Assert.AreEqual(HSSFColor.Gold.Index, borderFormatting.TopBorderColor);
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

            Assert.AreEqual(FontSuperScript.None, fontFormatting.EscapementType);
            Assert.AreEqual(-1, fontFormatting.FontColorIndex);
            Assert.AreEqual(-1, fontFormatting.FontHeight);
            Assert.AreEqual(0, fontFormatting.FontWeight);
            Assert.AreEqual(FontUnderlineType.None, fontFormatting.UnderlineType);

            fontFormatting.IsBold = (true);
            Assert.IsTrue(fontFormatting.IsBold);
            fontFormatting.IsBold = (false);
            Assert.IsFalse(fontFormatting.IsBold);

            fontFormatting.EscapementType = FontSuperScript.Sub;
            Assert.AreEqual(FontSuperScript.Sub, fontFormatting.EscapementType);
            fontFormatting.EscapementType = FontSuperScript.Super;
            Assert.AreEqual(FontSuperScript.Super, fontFormatting.EscapementType);
            fontFormatting.EscapementType = FontSuperScript.None;
            Assert.AreEqual(FontSuperScript.None, fontFormatting.EscapementType);

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

            fontFormatting.UnderlineType = FontUnderlineType.DoubleAccounting;
            Assert.AreEqual(FontUnderlineType.DoubleAccounting, fontFormatting.UnderlineType);

            fontFormatting.IsUnderlineTypeModified = (false);
            Assert.IsFalse(fontFormatting.IsUnderlineTypeModified);
            fontFormatting.IsUnderlineTypeModified = (true);
            Assert.IsTrue(fontFormatting.IsUnderlineTypeModified);
        }
        [Test]
        public void TestWrite() {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet();
            CFRuleRecord rr = CFRuleRecord.Create(sheet, (byte)ComparisonOperator.Between, "5", "10");

            PatternFormatting patternFormatting = new PatternFormatting();
            patternFormatting.FillPattern = FillPattern.Bricks;
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
        [Test]
        public void TestReserializeRefNTokens() 
        {
            
            RecordInputStream is1 = TestcaseRecordInputStream.Create (CFRuleRecord.sid, DATA_REFN);
            CFRuleRecord rr = new CFRuleRecord(is1);
            Ptg[] ptgs = rr.ParsedExpression1;
            Assert.AreEqual(3, ptgs.Length);
            if (ptgs[0] is RefPtg) {
                throw new AssertionException("Identified bug 45234");
            }
            Assert.AreEqual(typeof(RefNPtg), ptgs[0].GetType());
            RefNPtg refNPtg = (RefNPtg) ptgs[0];
            Assert.IsTrue(refNPtg.IsColRelative);
            Assert.IsTrue(refNPtg.IsRowRelative);
                
            byte[] data = rr.Serialize();

            TestcaseRecordInputStream.ConfirmRecordEncoding(CFRuleRecord.sid, DATA_REFN, data);
        }

        [Test]
        public void TestBug53691()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.CreateSheet() as HSSFSheet;

            CFRuleRecord record = CFRuleRecord.Create(sheet, (byte)ComparisonOperator.Between, "2", "5") as CFRuleRecord;

            CFRuleRecord clone = (CFRuleRecord)record.Clone();

            byte[] SerializedRecord = record.Serialize();
            byte[] SerializedClone = clone.Serialize();
            Assert.That(SerializedRecord, new EqualConstraint(SerializedClone));
        }

    }
}