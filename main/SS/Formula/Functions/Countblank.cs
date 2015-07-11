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

using NPOI.SS.Formula.Eval;
using System;
namespace NPOI.SS.Formula.Functions
{
    /**
     * Implementation for the function COUNTBLANK
     * <p>
     *  Syntax: COUNTBLANK ( range )
     *    <table border="0" cellpadding="1" cellspacing="0" summary="Parameter descriptions">
     *      <tr><th>range </th><td>is the range of cells to count blanks</td></tr>
     *    </table>
     * </p>
     *
     * @author Mads Mohr Christensen
     */
    public class Countblank : Fixed1ArgFunction
    {

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {

            double result;
            if (arg0 is RefEval)
            {
                result = CountUtils.CountMatchingCellsInRef((RefEval)arg0, predicate);
            }
            else if (arg0 is ThreeDEval)
            {
                result = CountUtils.CountMatchingCellsInArea((ThreeDEval)arg0, predicate);
            }
            else
            {
                throw new ArgumentException("Bad range arg type (" + arg0.GetType().Name + ")");
            }
            return new NumberEval(result);
        }

        private static IMatchPredicate predicate = new BlankPredicate();
        private class BlankPredicate : IMatchPredicate
        {
            #region I_MatchPredicate 成员

            public bool Matches(ValueEval valueEval)
            {
                return valueEval == BlankEval.instance;
            }

            #endregion
        }
    }
}