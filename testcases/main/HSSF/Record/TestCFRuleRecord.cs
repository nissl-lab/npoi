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
    using NUnit.Framework;using NUnit.Framework.Legacy;
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
            ClassicAssert.AreEqual(CFRuleRecord.CONDITION_TYPE_FORMULA, rule1.ConditionType);
            ClassicAssert.AreEqual((byte)ComparisonOperator.NoComparison, rule1.ComparisonOperation);
            ClassicAssert.IsNotNull(rule1.ParsedExpression1);
            ClassicAssert.AreSame(Ptg.EMPTY_PTG_ARRAY, rule1.ParsedExpression2);

            CFRuleRecord rule2 = CFRuleRecord.Create((HSSFSheet)sheet, (byte)ComparisonOperator.Between, "2", "5");
            ClassicAssert.AreEqual(CFRuleRecord.CONDITION_TYPE_CELL_VALUE_IS, rule2.ConditionType);
            ClassicAssert.AreEqual((byte)ComparisonOperator.Between, rule2.ComparisonOperation);
            ClassicAssert.IsNotNull(rule2.ParsedExpression1);
            ClassicAssert.IsNotNull(rule2.ParsedExpression2);

            CFRuleRecord rule3 = CFRuleRecord.Create((HSSFSheet)sheet, (byte)ComparisonOperator.Equal, null, null);
            ClassicAssert.AreEqual(CFRuleRecord.CONDITION_TYPE_CELL_VALUE_IS, rule3.ConditionType);
            ClassicAssert.AreEqual((byte)ComparisonOperator.Equal, rule3.ComparisonOperation);
            ClassicAssert.AreSame(Ptg.EMPTY_PTG_ARRAY, rule3.ParsedExpression2);
            ClassicAssert.AreSame(Ptg.EMPTY_PTG_ARRAY, rule3.ParsedExpression2);
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
            ClassicAssert.AreEqual(recordData.Length + 4, output.Length, "Output size"); //includes sid+recordlength

            for (int i = 0; i < recordData.Length; i++)
            {
                ClassicAssert.AreEqual(recordData[i], output[i + 4], "CFRuleRecord doesn't match");
            }
        }
        [Test]
        public void TestCreateCFRule12Record()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.CreateSheet() as HSSFSheet;
            CFRule12Record record = CFRule12Record.Create(sheet, "7");
            testCFRule12Record(record);
            // Serialize
            byte[] serializedRecord = record.Serialize();
            // Strip header
            byte[] recordData = new byte[serializedRecord.Length - 4];
            Array.Copy(serializedRecord, 4, recordData, 0, recordData.Length);
            // Deserialize
            record = new CFRule12Record(TestcaseRecordInputStream.Create(CFRule12Record.sid, recordData));
            // Serialize again
            byte[] output = record.Serialize();
            // Compare
            ClassicAssert.AreEqual(recordData.Length + 4, output.Length, "Output size"); //includes sid+recordlength
            for (int i = 0; i < recordData.Length; i++)
            {
                ClassicAssert.AreEqual(recordData[i], output[i + 4], "CFRule12Record doesn't match");
            }
        }

        [Test]
        public void TestCreateIconCFRule12Record()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.CreateSheet() as HSSFSheet;
            CFRule12Record record = CFRule12Record.Create(sheet, IconSet.GREY_5_ARROWS);
            record.MultiStateFormatting.Thresholds[1].Type = (byte)(RangeType.PERCENT.id);
            record.MultiStateFormatting.Thresholds[1].Value = (10d);
            record.MultiStateFormatting.Thresholds[2].Type = (byte)(RangeType.NUMBER.id);
            record.MultiStateFormatting.Thresholds[2].Value = (-4d);

            // Check it 
            testCFRule12Record(record);
            ClassicAssert.AreEqual(IconSet.GREY_5_ARROWS, record.MultiStateFormatting.IconSet);
            ClassicAssert.AreEqual(5, record.MultiStateFormatting.Thresholds.Length);
            // Serialize
            byte[] serializedRecord = record.Serialize();
            // Strip header
            byte[] recordData = new byte[serializedRecord.Length - 4];
            Array.Copy(serializedRecord, 4, recordData, 0, recordData.Length);
            // Deserialize
            record = new CFRule12Record(TestcaseRecordInputStream.Create(CFRule12Record.sid, recordData));

            // Check it has the icon, and the right number of thresholds
            ClassicAssert.AreEqual(IconSet.GREY_5_ARROWS, record.MultiStateFormatting.IconSet);
            ClassicAssert.AreEqual(5, record.MultiStateFormatting.Thresholds.Length);
            // Serialize again
            byte[] output = record.Serialize();
            // Compare
            ClassicAssert.AreEqual(recordData.Length + 4, output.Length, "Output size"); //includes sid+recordlength
            for (int i = 0; i < recordData.Length; i++)
            {
                ClassicAssert.AreEqual(recordData[i], output[i + 4], "CFRule12Record doesn't match");
            }
        }


        private void TestCFRuleRecord1(CFRuleRecord record)
        {
            testCFRuleBase(record);
            

            ClassicAssert.IsFalse(record.IsLeftBorderModified);
            record.IsLeftBorderModified = (true);
            ClassicAssert.IsTrue(record.IsLeftBorderModified);

            ClassicAssert.IsFalse(record.IsRightBorderModified);
            record.IsRightBorderModified = (true);
            ClassicAssert.IsTrue(record.IsRightBorderModified);

            ClassicAssert.IsFalse(record.IsTopBorderModified);
            record.IsTopBorderModified = (true);
            ClassicAssert.IsTrue(record.IsTopBorderModified);

            ClassicAssert.IsFalse(record.IsBottomBorderModified);
            record.IsBottomBorderModified = (true);
            ClassicAssert.IsTrue(record.IsBottomBorderModified);

            ClassicAssert.IsFalse(record.IsTopLeftBottomRightBorderModified);
            record.IsTopLeftBottomRightBorderModified = (true);
            ClassicAssert.IsTrue(record.IsTopLeftBottomRightBorderModified);

            ClassicAssert.IsFalse(record.IsBottomLeftTopRightBorderModified);
            record.IsBottomLeftTopRightBorderModified = (true);
            ClassicAssert.IsTrue(record.IsBottomLeftTopRightBorderModified);

            ClassicAssert.IsFalse(record.IsPatternBackgroundColorModified);
            record.IsPatternBackgroundColorModified = (true);
            ClassicAssert.IsTrue(record.IsPatternBackgroundColorModified);

            ClassicAssert.IsFalse(record.IsPatternColorModified);
            record.IsPatternColorModified = (true);
            ClassicAssert.IsTrue(record.IsPatternColorModified);

            ClassicAssert.IsFalse(record.IsPatternStyleModified);
            record.IsPatternStyleModified = (true);
            ClassicAssert.IsTrue(record.IsPatternStyleModified);
        }
        private void testCFRule12Record(CFRule12Record record)
        {
            ClassicAssert.AreEqual(CFRule12Record.sid, record.GetFutureRecordType());
            ClassicAssert.AreEqual("A1", record.GetAssociatedRange().FormatAsString());
            testCFRuleBase(record);
        }
        private void testCFRuleBase(CFRuleBase record)
        {
            FontFormatting fontFormatting = new FontFormatting();
            TestFontFormattingAccessors(fontFormatting);
            ClassicAssert.IsFalse(record.ContainsFontFormattingBlock);
            record.FontFormatting = (fontFormatting);
            ClassicAssert.IsTrue(record.ContainsFontFormattingBlock);

            BorderFormatting borderFormatting = new BorderFormatting();
            TestBorderFormattingAccessors(borderFormatting);
            ClassicAssert.IsFalse(record.ContainsBorderFormattingBlock);
            record.BorderFormatting = (borderFormatting);
            ClassicAssert.IsTrue(record.ContainsBorderFormattingBlock);

            PatternFormatting patternFormatting = new PatternFormatting();
            TestPatternFormattingAccessors(patternFormatting);
            ClassicAssert.IsFalse(record.ContainsPatternFormattingBlock);
            record.PatternFormatting = (patternFormatting);
            ClassicAssert.IsTrue(record.ContainsPatternFormattingBlock);
        }

        private void TestPatternFormattingAccessors(PatternFormatting patternFormatting)
        {
            patternFormatting.FillBackgroundColor = (HSSFColor.Green.Index);
            ClassicAssert.AreEqual(HSSFColor.Green.Index, patternFormatting.FillBackgroundColor);

            patternFormatting.FillForegroundColor = (HSSFColor.Indigo.Index);
            ClassicAssert.AreEqual(HSSFColor.Indigo.Index, patternFormatting.FillForegroundColor);

            patternFormatting.FillPattern = FillPattern.Diamonds;
            ClassicAssert.AreEqual(FillPattern.Diamonds, patternFormatting.FillPattern);
        }

        private void TestBorderFormattingAccessors(BorderFormatting borderFormatting)
        {
            borderFormatting.IsBackwardDiagonalOn = (false);
            ClassicAssert.IsFalse(borderFormatting.IsBackwardDiagonalOn);
            borderFormatting.IsBackwardDiagonalOn = (true);
            ClassicAssert.IsTrue(borderFormatting.IsBackwardDiagonalOn);

            borderFormatting.BorderBottom = BorderStyle.Dotted;
            ClassicAssert.AreEqual(BorderStyle.Dotted, borderFormatting.BorderBottom);

            borderFormatting.BorderDiagonal = (BorderStyle.Medium);
            ClassicAssert.AreEqual(BorderStyle.Medium, borderFormatting.BorderDiagonal);

            borderFormatting.BorderLeft = (BorderStyle.MediumDashDotDot);
            ClassicAssert.AreEqual(BorderStyle.MediumDashDotDot, borderFormatting.BorderLeft);

            borderFormatting.BorderRight = (BorderStyle.MediumDashed);
            ClassicAssert.AreEqual(BorderStyle.MediumDashed, borderFormatting.BorderRight);

            borderFormatting.BorderTop = (BorderStyle.Hair);
            ClassicAssert.AreEqual(BorderStyle.Hair, borderFormatting.BorderTop);

            borderFormatting.BottomBorderColor = (HSSFColor.Aqua.Index);
            ClassicAssert.AreEqual(HSSFColor.Aqua.Index, borderFormatting.BottomBorderColor);

            borderFormatting.DiagonalBorderColor = (HSSFColor.Red.Index);
            ClassicAssert.AreEqual(HSSFColor.Red.Index, borderFormatting.DiagonalBorderColor);

            ClassicAssert.IsFalse(borderFormatting.IsForwardDiagonalOn);
            borderFormatting.IsForwardDiagonalOn = (true);
            ClassicAssert.IsTrue(borderFormatting.IsForwardDiagonalOn);

            borderFormatting.LeftBorderColor = (HSSFColor.Black.Index);
            ClassicAssert.AreEqual(HSSFColor.Black.Index, borderFormatting.LeftBorderColor);

            borderFormatting.RightBorderColor = (HSSFColor.Blue.Index);
            ClassicAssert.AreEqual(HSSFColor.Blue.Index, borderFormatting.RightBorderColor);

            borderFormatting.TopBorderColor = (HSSFColor.Gold.Index);
            ClassicAssert.AreEqual(HSSFColor.Gold.Index, borderFormatting.TopBorderColor);
        }


        private void TestFontFormattingAccessors(FontFormatting fontFormatting)
        {
            // Check for defaults
            ClassicAssert.IsFalse(fontFormatting.IsEscapementTypeModified);
            ClassicAssert.IsFalse(fontFormatting.IsFontCancellationModified);
            ClassicAssert.IsFalse(fontFormatting.IsFontOutlineModified);
            ClassicAssert.IsFalse(fontFormatting.IsFontShadowModified);
            ClassicAssert.IsFalse(fontFormatting.IsFontStyleModified);
            ClassicAssert.IsFalse(fontFormatting.IsUnderlineTypeModified);
            ClassicAssert.IsFalse(fontFormatting.IsFontWeightModified);

            ClassicAssert.IsFalse(fontFormatting.IsBold);
            ClassicAssert.IsFalse(fontFormatting.IsItalic);
            ClassicAssert.IsFalse(fontFormatting.IsOutlineOn);
            ClassicAssert.IsFalse(fontFormatting.IsShadowOn);
            ClassicAssert.IsFalse(fontFormatting.IsStruckout);

            ClassicAssert.AreEqual(FontSuperScript.None, fontFormatting.EscapementType);
            ClassicAssert.AreEqual(-1, fontFormatting.FontColorIndex);
            ClassicAssert.AreEqual(-1, fontFormatting.FontHeight);
            ClassicAssert.AreEqual(0, fontFormatting.FontWeight);
            ClassicAssert.AreEqual(FontUnderlineType.None, fontFormatting.UnderlineType);

            fontFormatting.IsBold = (true);
            ClassicAssert.IsTrue(fontFormatting.IsBold);
            fontFormatting.IsBold = (false);
            ClassicAssert.IsFalse(fontFormatting.IsBold);

            fontFormatting.EscapementType = FontSuperScript.Sub;
            ClassicAssert.AreEqual(FontSuperScript.Sub, fontFormatting.EscapementType);
            fontFormatting.EscapementType = FontSuperScript.Super;
            ClassicAssert.AreEqual(FontSuperScript.Super, fontFormatting.EscapementType);
            fontFormatting.EscapementType = FontSuperScript.None;
            ClassicAssert.AreEqual(FontSuperScript.None, fontFormatting.EscapementType);

            fontFormatting.IsEscapementTypeModified = (false);
            ClassicAssert.IsFalse(fontFormatting.IsEscapementTypeModified);
            fontFormatting.IsEscapementTypeModified = (true);
            ClassicAssert.IsTrue(fontFormatting.IsEscapementTypeModified);

            fontFormatting.IsFontWeightModified = (false);
            ClassicAssert.IsFalse(fontFormatting.IsFontWeightModified);
            fontFormatting.IsFontWeightModified = (true);
            ClassicAssert.IsTrue(fontFormatting.IsFontWeightModified);

            fontFormatting.IsFontCancellationModified = (false);
            ClassicAssert.IsFalse(fontFormatting.IsFontCancellationModified);
            fontFormatting.IsFontCancellationModified = (true);
            ClassicAssert.IsTrue(fontFormatting.IsFontCancellationModified);

            fontFormatting.FontColorIndex = ((short)10);
            ClassicAssert.AreEqual(10, fontFormatting.FontColorIndex);

            fontFormatting.FontHeight = (100);
            ClassicAssert.AreEqual(100, fontFormatting.FontHeight);

            fontFormatting.IsFontOutlineModified = (false);
            ClassicAssert.IsFalse(fontFormatting.IsFontOutlineModified);
            fontFormatting.IsFontOutlineModified = (true);
            ClassicAssert.IsTrue(fontFormatting.IsFontOutlineModified);

            fontFormatting.IsFontShadowModified = (false);
            ClassicAssert.IsFalse(fontFormatting.IsFontShadowModified);
            fontFormatting.IsFontShadowModified = (true);
            ClassicAssert.IsTrue(fontFormatting.IsFontShadowModified);

            fontFormatting.IsFontStyleModified = (false);
            ClassicAssert.IsFalse(fontFormatting.IsFontStyleModified);
            fontFormatting.IsFontStyleModified = (true);
            ClassicAssert.IsTrue(fontFormatting.IsFontStyleModified);

            fontFormatting.IsItalic = (false);
            ClassicAssert.IsFalse(fontFormatting.IsItalic);
            fontFormatting.IsItalic = (true);
            ClassicAssert.IsTrue(fontFormatting.IsItalic);

            fontFormatting.IsOutlineOn = (false);
            ClassicAssert.IsFalse(fontFormatting.IsOutlineOn);
            fontFormatting.IsOutlineOn = (true);
            ClassicAssert.IsTrue(fontFormatting.IsOutlineOn);

            fontFormatting.IsShadowOn = (false);
            ClassicAssert.IsFalse(fontFormatting.IsShadowOn);
            fontFormatting.IsShadowOn = (true);
            ClassicAssert.IsTrue(fontFormatting.IsShadowOn);

            fontFormatting.IsStruckout = (false);
            ClassicAssert.IsFalse(fontFormatting.IsStruckout);
            fontFormatting.IsStruckout = (true);
            ClassicAssert.IsTrue(fontFormatting.IsStruckout);

            fontFormatting.UnderlineType = FontUnderlineType.DoubleAccounting;
            ClassicAssert.AreEqual(FontUnderlineType.DoubleAccounting, fontFormatting.UnderlineType);

            fontFormatting.IsUnderlineTypeModified = (false);
            ClassicAssert.IsFalse(fontFormatting.IsUnderlineTypeModified);
            fontFormatting.IsUnderlineTypeModified = (true);
            ClassicAssert.IsTrue(fontFormatting.IsUnderlineTypeModified);
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
            ClassicAssert.AreEqual(26, data.Length);
            ClassicAssert.AreEqual(3, LittleEndian.GetShort(data, 6));
            ClassicAssert.AreEqual(3, LittleEndian.GetShort(data, 8));

            int flags = LittleEndian.GetInt(data, 10);
            ClassicAssert.AreEqual(0x00380000, flags & 0x00380000,"unused flags should be 111");
            ClassicAssert.AreEqual(0, flags & 0x03C00000,"undocumented flags should be 0000"); // Otherwise Excel s unhappy
            // check all remaining flag bits (some are not well understood yet)
            ClassicAssert.AreEqual(0x203FFFFF, flags);
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
            ClassicAssert.AreEqual(3, ptgs.Length);
            if (ptgs[0] is RefPtg) {
                throw new AssertionException("Identified bug 45234");
            }
            ClassicAssert.AreEqual(typeof(RefNPtg), ptgs[0].GetType());
            RefNPtg refNPtg = (RefNPtg) ptgs[0];
            ClassicAssert.IsTrue(refNPtg.IsColRelative);
            ClassicAssert.IsTrue(refNPtg.IsRowRelative);
                
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
        [Test]
        public void TestBug57231_rewrite()
        {
            HSSFWorkbook wb = HSSFITestDataProvider.Instance.OpenSampleWorkbook("57231_MixedGasReport.xls") as HSSFWorkbook;
            ClassicAssert.AreEqual(7, wb.NumberOfSheets);
            wb = HSSFITestDataProvider.Instance.WriteOutAndReadBack(wb) as HSSFWorkbook;
            ClassicAssert.AreEqual(7, wb.NumberOfSheets);
        }
    }
}