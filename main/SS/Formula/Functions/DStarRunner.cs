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

namespace NPOI.SS.Formula.Functions
{
    using System;

    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Util;
    using NPOI.Util;

    /**
     * This class performs a D* calculation. It takes an {@link IDStarAlgorithm} object and
     * uses it for calculating the result value. Iterating a database and Checking the
     * entries against the Set of conditions is done here.
     */
    public class DStarRunner : Function3Arg
    {

        public enum DStarAlgorithmEnum
        {
            DGET,
            DMIN,
            // DMAX, // DMAX is not yet implemented
        }

        private DStarAlgorithmEnum algoType;

        public DStarRunner(DStarAlgorithmEnum algorithm)
        {
            this.algoType = algorithm;
        }

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length == 3)
            {
                return Evaluate(srcRowIndex, srcColumnIndex, args[0], args[1], args[2]);
            }
            else
            {
                return ErrorEval.VALUE_INVALID;
            }
        }

        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex,
                ValueEval database, ValueEval filterColumn, ValueEval conditionDatabase)
        {
            // Input Processing and error Checks.
            if (!(database is AreaEval) || !(conditionDatabase is AreaEval))
            {
                return ErrorEval.VALUE_INVALID;
            }
            AreaEval db = (AreaEval)database;
            AreaEval cdb = (AreaEval)conditionDatabase;

            try
            {
                filterColumn = OperandResolver.GetSingleValue(filterColumn, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            int fc;
            try
            {
                fc = GetColumnForName(filterColumn, db);
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }
            if (fc == -1)
            { // column not found
                return ErrorEval.VALUE_INVALID;
            }

            // Create an algorithm runner.
            IDStarAlgorithm algorithm = null;
            switch (algoType)
            {
                case DStarAlgorithmEnum.DGET: algorithm = new DGet(); break;
                case DStarAlgorithmEnum.DMIN: algorithm = new DMin(); break;
                default:
                    throw new InvalidOperationException("Unexpected algorithm type " + algoType + " encountered.");
            }

            // Iterate over all db entries.
            int height = db.Height;
            for (int row = 1; row < height; ++row)
            {
                bool matches = true;
                try
                {
                    matches = FullFillsConditions(db, row, cdb);
                }
                catch (EvaluationException)
                {
                    return ErrorEval.VALUE_INVALID;
                }
                // Filter each entry.
                if (matches)
                {
                    ValueEval currentValueEval = ResolveReference(db, row, fc);
                    // Pass the match to the algorithm and conditionally abort the search.
                    bool shouldContinue = algorithm.ProcessMatch(currentValueEval);
                    if (!shouldContinue)
                    {
                        break;
                    }
                }
            }

            // Return the result of the algorithm.
            return algorithm.Result;
        }

        private enum Operator
        {
            largerThan,
            largerEqualThan,
            smallerThan,
            smallerEqualThan,
            equal
        }

        /**
         * 
         *
         * @param nameValueEval Must not be a RefEval or AreaEval. Thus make sure resolveReference() is called on the value first!
         * @param db
         * @return
         * @throws EvaluationException
         */
        private static int GetColumnForName(ValueEval nameValueEval, AreaEval db)
        {
            String name = OperandResolver.CoerceValueToString(nameValueEval);
            return GetColumnForString(db, name);
        }

        /**
         * For a given database returns the column number for a column heading.
         *
         * @param db Database.
         * @param name Column heading.
         * @return Corresponding column number.
         * @If it's not possible to turn all headings into strings.
         */
        private static int GetColumnForString(AreaEval db, String name)
        {
            int resultColumn = -1;
            int width = db.Width;
            for (int column = 0; column < width; ++column)
            {
                ValueEval columnNameValueEval = ResolveReference(db, 0, column);
                if (columnNameValueEval is BlankEval)
                {
                    continue;
                }
                if (columnNameValueEval is ErrorEval)
                {
                    continue;
                }
                String columnName = OperandResolver.CoerceValueToString(columnNameValueEval);
                if (name.Equals(columnName))
                {
                    resultColumn = column;
                    break;
                }
            }
            return resultColumn;
        }

        /**
         * Checks a row in a database against a condition database.
         *
         * @param db Database.
         * @param row The row in the database to Check.
         * @param cdb The condition database to use for Checking.
         * @return Whether the row matches the conditions.
         * @If references could not be Resolved or comparison
         * operators and operands didn't match.
         */
        private static bool FullFillsConditions(AreaEval db, int row, AreaEval cdb)
        {
            // Only one row must match to accept the input, so rows are ORed.
            // Each row is made up of cells where each cell is a condition,
            // all have to match, so they are ANDed.
            int height = cdb.Height;
            for (int conditionRow = 1; conditionRow < height; ++conditionRow)
            {
                bool matches = true;
                int width = cdb.Width;
                for (int column = 0; column < width; ++column)  // columns are ANDed
                { 
                    // Whether the condition column matches a database column, if not it's a
                    // special column that accepts formulas.
                    bool columnCondition = true;
                    ValueEval condition = null;

                    // The condition to apply.
                    condition = ResolveReference(cdb, conditionRow, column);

                    // If the condition is empty it matches.
                    if (condition is BlankEval)
                        continue;
                    // The column in the DB to apply the condition to.
                    ValueEval targetHeader = ResolveReference(cdb, 0, column);

                    if (!(targetHeader is StringValueEval))
                    {
                        throw new EvaluationException(ErrorEval.VALUE_INVALID);
                    }
                        
                    if (GetColumnForName(targetHeader, db) == -1)
                        // No column found, it's again a special column that accepts formulas.
                        columnCondition = false;

                    if (columnCondition == true)
                    { // normal column condition
                        // Should not throw, Checked above.
                        ValueEval value = ResolveReference(db, row, GetColumnForName(targetHeader, db));
                        if (!testNormalCondition(value, condition))
                        { 
                            matches = false;
                            break;
                        }
                    }
                    else
                    { // It's a special formula condition.
                      // TODO: Check whether the condition cell contains a formula and return #VALUE! if it doesn't.
                        if (string.IsNullOrEmpty(OperandResolver.CoerceValueToString(condition)))
                        {
                            throw new EvaluationException(ErrorEval.VALUE_INVALID);
                        }
                        throw new NotImplementedException(
                                "D* function with formula conditions");
                    }
                }
                if (matches == true)
                {
                    return true;
                }
            }
            return false;
        }

        /**
         * Test a value against a simple (&lt; &gt; &lt;= &gt;= = starts-with) condition string.
         *
         * @param value The value to Check.
         * @param condition The condition to check for.
         * @return Whether the condition holds.
         * @If comparison operator and operands don't match.
         */
        private static bool testNormalCondition(ValueEval value, ValueEval condition)
        {
            if (condition is StringEval) {
                String conditionString = ((StringEval)condition).StringValue;

                if (conditionString.StartsWith("<"))
                { // It's a </<= condition.
                    String number = conditionString.Substring(1);
                    if (number.StartsWith("="))
                    {
                        number = number.Substring(1);
                        return testNumericCondition(value, Operator.smallerEqualThan, number);
                    }
                    else
                    {
                        return testNumericCondition(value, Operator.smallerThan, number);
                    }
                }
                else if (conditionString.StartsWith(">"))
                { // It's a >/>= condition.
                    String number = conditionString.Substring(1);
                    if (number.StartsWith("="))
                    {
                        number = number.Substring(1);
                        return testNumericCondition(value, Operator.largerEqualThan, number);
                    }
                    else
                    {
                        return testNumericCondition(value, Operator.largerThan, number);
                    }
                }
                else if (conditionString.StartsWith("="))
                { // It's a = condition.
                    String stringOrNumber = conditionString.Substring(1);

                    if (string.IsNullOrEmpty(stringOrNumber))
                    {
                        return value is BlankEval;
                    }
                    // Distinguish between string and number.
                    bool itsANumber = false;
                    try
                    {
                        int.Parse(stringOrNumber);
                        itsANumber = true;
                    }
                    catch (FormatException)
                    { // It's not an int.
                        try
                        {
                            Double.Parse(stringOrNumber);
                            itsANumber = true;
                        }
                        catch (FormatException)
                        { // It's a string.
                            itsANumber = false;
                        }
                    }
                    if (itsANumber)
                    {
                        return testNumericCondition(value, Operator.equal, stringOrNumber);
                    }
                    else
                    { // It's a string.
                        String valueString = value is BlankEval ? "" : OperandResolver.CoerceValueToString(value);
                        return stringOrNumber.Equals(valueString);
                    }
                }
                else
                { // It's a text starts-with condition.
                    if (string.IsNullOrEmpty(conditionString))
                    {
                        return value is StringEval;
                    }
                    else
                    {
                        String valueString = value is BlankEval ? "" : OperandResolver.CoerceValueToString(value);
                        return valueString.StartsWith(conditionString);
                    }
                }
            }
            else if (condition is NumericValueEval) {
                double conditionNumber = ((NumericValueEval)condition).NumberValue;
                Double? valueNumber = GetNumberFromValueEval(value);
                if (valueNumber == null)
                {
                    return false;
                }

                return conditionNumber == valueNumber;
            }
            else if (condition is ErrorEval) {
                if (value is ErrorEval) {
                    return ((ErrorEval)condition).ErrorCode == ((ErrorEval)value).ErrorCode;
                }
                else {
                    return false;
                }
            }
            else {
                return false;
            }
            
        }

        /**
         * Test whether a value matches a numeric condition.
         * @param valueEval Value to Check.
         * @param op Comparator to use.
         * @param condition Value to check against.
         * @return whether the condition holds.
         * @If it's impossible to turn the condition into a number.
         */
        private static bool testNumericCondition(
                ValueEval valueEval, Operator op, String condition)
        {
            // Construct double from ValueEval.
            if (!(valueEval is NumericValueEval))
                return false;
            double value = ((NumericValueEval)valueEval).NumberValue;

            // Construct double from condition.
            double conditionValue = 0.0;
            try
            {
                int intValue = Int32.Parse(condition);
                conditionValue = intValue;
            }
            catch (FormatException)
            { // It's not an int.
                try
                {
                    conditionValue = Double.Parse(condition);
                }
                catch (FormatException)
                { // It's not a double.
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
            }

            int result = NumberComparer.Compare(value, conditionValue);
            switch (op)
            {
                case Operator.largerThan:
                    return result > 0;
                case Operator.largerEqualThan:
                    return result >= 0;
                case Operator.smallerThan:
                    return result < 0;
                case Operator.smallerEqualThan:
                    return result <= 0;
                case Operator.equal:
                    return result == 0;
            }
            return false; // Can not be reached.
        }

        private static Double? GetNumberFromValueEval(ValueEval value)
        {
            if (value is NumericValueEval)
            {
                return ((NumericValueEval)value).NumberValue;
            }
            else if (value is StringValueEval)
            {
                String stringValue = ((StringValueEval)value).StringValue;
                try
                {
                    return Double.Parse(stringValue);
                }
                catch (FormatException)
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }
        /**
         * Resolve a ValueEval that's in an AreaEval.
         *
         * @param db AreaEval from which the cell to resolve is retrieved. 
         * @param dbRow Relative row in the AreaEval.
         * @param dbCol Relative column in the AreaEval.
         * @return A ValueEval that is a NumberEval, StringEval, BoolEval, BlankEval or ErrorEval.
         */
        private static ValueEval ResolveReference(AreaEval db, int dbRow, int dbCol)
        {
            try
            {
                return OperandResolver.GetSingleValue(db.GetValue(dbRow, dbCol), db.FirstRow + dbRow, db.FirstColumn + dbCol);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
    }

}