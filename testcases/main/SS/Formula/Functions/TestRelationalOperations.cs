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

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    [TestFixture]
    public class TestRelationalOperations
    {

        /// <summary>
        /// <para>
        ///  (1, 1)(1, 1) = 1
        /// </para>
        /// <para>
        ///   evaluates to
        /// </para>
        /// <para>
        ///   (TRUE, TRUE)(TRUE, TRUE)
        /// </para>
        /// </summary>
        [Test]
        public void TstEqMatrixByScalar_Numbers()
        {
            ValueEval[] values = new ValueEval[4];
            for(int i = 0; i < values.Length; i++)
            {
                values[i] = new NumberEval(1);
            }

            ValueEval arg1 = EvalFactory.CreateAreaEval("A1:B2", values);
            ValueEval arg2 = EvalFactory.CreateRefEval("D1", new NumberEval(1));

            RelationalOperationEval eq = (RelationalOperationEval)RelationalOperationEval.EqualEval;
            ValueEval result = eq.EvaluateArray(new ValueEval[]{ arg1, arg2}, 2, 5);

            ClassicAssert.AreEqual(typeof(CacheAreaEval), result.GetType(), "expected CacheAreaEval");
            CacheAreaEval ce = (CacheAreaEval)result;
            ClassicAssert.AreEqual(2, ce.Width);
            ClassicAssert.AreEqual(2, ce.Height);
            for(int i = 0; i < ce.Height; i++)
            {
                for(int j = 0; j < ce.Width; j++)
                {
                    ClassicAssert.AreEqual(BoolEval.TRUE, ce.GetRelativeValue(i, j));
                }
            }
        }
        [Test]
        public void TestEqMatrixByScalar_String()
        {
            ValueEval[] values = new ValueEval[4];
            for(int i = 0; i < values.Length; i++)
            {
                values[i] = new StringEval("ABC");
            }

            ValueEval arg1 = EvalFactory.CreateAreaEval("A1:B2", values);
            ValueEval arg2 = EvalFactory.CreateRefEval("D1", new StringEval("ABC"));
            RelationalOperationEval eq = (RelationalOperationEval)RelationalOperationEval.EqualEval;
            ValueEval result = eq.EvaluateArray(new ValueEval[]{ arg1, arg2}, 2, 5);

            ClassicAssert.AreEqual(typeof(CacheAreaEval), result.GetType(), "expected CacheAreaEval");
            CacheAreaEval ce = (CacheAreaEval)result;
            ClassicAssert.AreEqual(2, ce.Width);
            ClassicAssert.AreEqual(2, ce.Height);
            for(int i = 0; i < ce.Height; i++)
            {
                for(int j = 0; j < ce.Width; j++)
                {
                    ClassicAssert.AreEqual(BoolEval.TRUE, ce.GetRelativeValue(i, j));
                }
            }
        }
        [Test]
        public void TestEqMatrixBy_Row()
        {
            ValueEval[] matrix = {
                new NumberEval(-1), new NumberEval(1),
                new NumberEval(-1), new NumberEval(1)
            };


            ValueEval[] row = {
                new NumberEval(1), new NumberEval(1), new NumberEval(1)
            };

            ValueEval[] expected = {
                BoolEval.FALSE, BoolEval.TRUE, ErrorEval.VALUE_INVALID,
                BoolEval.FALSE, BoolEval.TRUE, ErrorEval.VALUE_INVALID
            };

            ValueEval arg1 = EvalFactory.CreateAreaEval("A1:B2", matrix);
            ValueEval arg2 = EvalFactory.CreateAreaEval("A4:C4", row);
            RelationalOperationEval eq = (RelationalOperationEval)RelationalOperationEval.EqualEval;
            ValueEval result = eq.EvaluateArray(new ValueEval[]{ arg1, arg2}, 4, 5);

            ClassicAssert.AreEqual(typeof(CacheAreaEval), result.GetType(), "expected CacheAreaEval");
            CacheAreaEval ce = (CacheAreaEval)result;
            ClassicAssert.AreEqual(3, ce.Width);
            ClassicAssert.AreEqual(2, ce.Height);
            int idx = 0;
            for(int i = 0; i < ce.Height; i++)
            {
                for(int j = 0; j < ce.Width; j++)
                {
                    ClassicAssert.AreEqual(expected[idx++], ce.GetRelativeValue(i, j), "[" + i + "," + j + "]");
                }
            }
        }
        [Test]
        public void TestEqMatrixBy_Column()
        {
            ValueEval[] matrix = {
                new NumberEval(-1), new NumberEval(1),
                new NumberEval(-1), new NumberEval(1)
            };

            ValueEval[] column = {
                new NumberEval(1),
                new NumberEval(1),
                new NumberEval(1)
            };

            ValueEval[] expected = {
                BoolEval.FALSE, BoolEval.TRUE,
                BoolEval.FALSE, BoolEval.TRUE,
                ErrorEval.VALUE_INVALID, ErrorEval.VALUE_INVALID
            };

            ValueEval arg1 = EvalFactory.CreateAreaEval("A1:B2", matrix);
            ValueEval arg2 = EvalFactory.CreateAreaEval("A6:A8", column);
            RelationalOperationEval eq = (RelationalOperationEval)RelationalOperationEval.EqualEval;
            ValueEval result = eq.EvaluateArray(new ValueEval[]{ arg1, arg2}, 4, 6);

            ClassicAssert.AreEqual(typeof(CacheAreaEval), result.GetType(), "expected CacheAreaEval");
            CacheAreaEval ce = (CacheAreaEval)result;
            ClassicAssert.AreEqual(2, ce.Width);
            ClassicAssert.AreEqual(3, ce.Height);
            int idx = 0;
            for(int i = 0; i < ce.Height; i++)
            {
                for(int j = 0; j < ce.Width; j++)
                {
                    ClassicAssert.AreEqual(expected[idx++], ce.GetRelativeValue(i, j), "[" + i + "," + j + "]");
                }
            }
        }
        [Test]
        public void TestEqMatrixBy_Matrix()
        {
            // A1:B2
            ValueEval[] matrix1 = {
                new NumberEval(-1), new NumberEval(1),
                new NumberEval(-1), new NumberEval(1)
            };

            // A10:C12
            ValueEval[] matrix2 = {
                new NumberEval(1), new NumberEval(1), new NumberEval(1),
                new NumberEval(1), new NumberEval(1), new NumberEval(1),
                new NumberEval(1), new NumberEval(1), new NumberEval(1)
            };

            ValueEval[] expected = {
                BoolEval.FALSE, BoolEval.TRUE, ErrorEval.VALUE_INVALID,
                BoolEval.FALSE, BoolEval.TRUE, ErrorEval.VALUE_INVALID,
                ErrorEval.VALUE_INVALID, ErrorEval.VALUE_INVALID, ErrorEval.VALUE_INVALID
            };

            ValueEval arg1 = EvalFactory.CreateAreaEval("A1:B2", matrix1);
            ValueEval arg2 = EvalFactory.CreateAreaEval("A10:C12", matrix2);
            RelationalOperationEval eq = (RelationalOperationEval)RelationalOperationEval.EqualEval;
            ValueEval result = eq.EvaluateArray(new ValueEval[]{ arg1, arg2}, 4, 6);

            ClassicAssert.AreEqual(typeof(CacheAreaEval), result.GetType(), "expected CacheAreaEval");
            CacheAreaEval ce = (CacheAreaEval)result;
            ClassicAssert.AreEqual(3, ce.Width);
            ClassicAssert.AreEqual(3, ce.Height);
            int idx = 0;
            for(int i = 0; i < ce.Height; i++)
            {
                for(int j = 0; j < ce.Width; j++)
                {
                    ClassicAssert.AreEqual(expected[idx++], ce.GetRelativeValue(i, j), "[" + i + "," + j + "]");
                }
            }
        }

    }
}


