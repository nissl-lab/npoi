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
namespace NPOI.SS.Formula;

using junit.framework.TestCase;
using NPOI.SS.ITestDataProvider;
using NPOI.SS.Formula.Eval.ErrorEval;
using NPOI.SS.Formula.Eval.StringEval;
using NPOI.SS.Formula.Eval.ValueEval;
using NPOI.SS.Formula.functions.FreeRefFunction;
using NPOI.SS.Formula.udf.DefaultUDFFinder;
using NPOI.SS.Formula.udf.UDFFinder;
using NPOI.SS.UserModel.Cell;
using NPOI.SS.UserModel.FormulaEvaluator;
using NPOI.SS.UserModel.Sheet;
using NPOI.SS.UserModel.Workbook;

/**
 * Test Setting / Evaluating of Analysis Toolpack and user-defined functions
 *
 * @author Yegor Kozlov
 */
public class BaseTestExternalFunctions  {
    // define two custom user-defined functions
    private static class MyFunc : FreeRefFunction {
        public MyFunc() {
            //
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec) {
            if (args.Length != 1 || !(args[0] is StringEval)) {
                return ErrorEval.VALUE_INVALID;
            }
            StringEval input = (StringEval) args[0];
            return new StringEval(input.StringValue + "abc");
        }
    }

    private static class MyFunc2 : FreeRefFunction {
        public MyFunc2() {
            //
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec) {
            if (args.Length != 1 || !(args[0] is StringEval)) {
                return ErrorEval.VALUE_INVALID;
            }
            StringEval input = (StringEval) args[0];
            return new StringEval(input.StringValue + "abc2");
        }
    }

    /**
     * register the two Test UDFs in a UDF Finder, to be passed to the workbook
     */
    private static UDFFinder customToolpack = new DefaultUDFFinder(
            new String[] { "myFunc", "myFunc2"},
            new FreeRefFunction[] { new MyFunc(), new MyFunc2()}
    );


    protected ITestDataProvider _testDataProvider;

    /**
     * @param TestDataProvider an object that provides Test data in HSSF / XSSF specific way
     */
    protected BaseTestExternalFunctions(ITestDataProvider TestDataProvider) {
        _testDataProvider = TestDataProvider;
    }

    public void TestExternalFunctions() {
        Workbook wb = _testDataProvider.CreateWorkbook();

        Sheet sh = wb.CreateSheet();

        Cell cell1 = sh.CreateRow(0).createCell(0);
        cell1.SetCellFormula("ISODD(1)+ISEVEN(2)"); // functions from the Excel Analysis Toolpack
        Assert.AreEqual("ISODD(1)+ISEVEN(2)", cell1.GetCellFormula());

        Cell cell2 = sh.CreateRow(1).createCell(0);
        try {
            cell2.SetCellFormula("MYFUNC(\"B1\")");
            Assert.Fail("Should fail because MYFUNC is an unknown function");
        } catch (FormulaParseException e){
            ; //expected
        }

        wb.AddToolPack(customToolpack);

        cell2.SetCellFormula("MYFUNC(\"B1\")");
        Assert.AreEqual("MYFUNC(\"B1\")", cell2.GetCellFormula());

        Cell cell3 = sh.CreateRow(2).createCell(0);
        cell3.SetCellFormula("MYFUNC2(\"C1\")&\"-\"&A2");  //where A2 is defined above
        Assert.AreEqual("MYFUNC2(\"C1\")&\"-\"&A2", cell3.GetCellFormula());

        FormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
        Assert.AreEqual(2.0, Evaluator.evaluate(cell1).GetNumberValue());
        Assert.AreEqual("B1abc", Evaluator.evaluate(cell2).StringValue);
        Assert.AreEqual("C1abc2-B1abc", Evaluator.evaluate(cell3).StringValue);

    }

    /**
     * Test invoking saved ATP functions
     *
     * @param TestFile  either atp.xls or atp.xlsx
     */
    public void baseTestInvokeATP(String TestFile){
        Workbook wb = _testDataProvider.OpenSampleWorkbook(testFile);
        FormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

        Sheet sh  = wb.GetSheetAt(0);
        // these two are not imlemented in r
        Assert.AreEqual("DELTA(1.3,1.5)", sh.GetRow(0).GetCell(1).GetCellFormula());
        Assert.AreEqual("COMPLEX(2,4)", sh.GetRow(1).GetCell(1).GetCellFormula());

        Cell cell2 = sh.GetRow(2).GetCell(1);
        Assert.AreEqual("ISODD(2)", cell2.GetCellFormula());
        Assert.AreEqual(false, Evaluator.evaluate(cell2).GetBooleanValue());
        Assert.AreEqual(Cell.CELL_TYPE_BOOLEAN, Evaluator.evaluateFormulaCell(cell2));

        Cell cell3 = sh.GetRow(3).GetCell(1);
        Assert.AreEqual("ISEVEN(2)", cell3.GetCellFormula());
        Assert.AreEqual(true, Evaluator.evaluate(cell3).GetBooleanValue());
        Assert.AreEqual(Cell.CELL_TYPE_BOOLEAN, Evaluator.evaluateFormulaCell(cell3));

    }

}

