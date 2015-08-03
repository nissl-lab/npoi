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

    /**
     * This class performs a D* calculation. It takes an {@link IDStarAlgorithm} object and
     * uses it for calculating the result value. Iterating a database and Checking the
     * entries against the Set of conditions is done here.
     */
    public class DStarRunner : Function3Arg
    {
        private IDStarAlgorithm algorithm;

        public DStarRunner(IDStarAlgorithm algorithm)
        {
            this.algorithm = algorithm;
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
            if (!(database is TwoDEval) || !(conditionDatabase is TwoDEval))
            {
                return ErrorEval.VALUE_INVALID;
            }
            TwoDEval db = (TwoDEval)database;
            TwoDEval cdb = (TwoDEval)conditionDatabase;

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

            // Reset algorithm.
            algorithm.Reset();

            // Iterate over all db entries.
            for (int row = 1; row < db.Height; ++row)
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
                    try
                    {
                        ValueEval currentValueEval = solveReference(db.GetValue(row, fc));
                        // Pass the match to the algorithm and conditionally abort the search.
                        bool shouldContinue = algorithm.ProcessMatch(currentValueEval);
                        if (!shouldContinue)
                        {
                            break;
                        }
                    }
                    catch (EvaluationException e)
                    {
                        return e.GetErrorEval();
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
         * Resolve reference(-chains) until we have a normal value.
         *
         * @param field a ValueEval which can be a RefEval.
         * @return a ValueEval which is guaranteed not to be a RefEval
         * @If a multi-sheet reference was found along the way.
         */
        private static ValueEval solveReference(ValueEval field)
        {
            if (field is RefEval)
            {
                RefEval refEval = (RefEval)field;
                if (refEval.NumberOfSheets > 1)
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                return solveReference(refEval.GetInnerValueEval(refEval.FirstSheetIndex));
            }
            else
            {
                return field;
            }
        }

        /**
         * Returns the first column index that matches the given name. The name can either be
         * a string or an integer, when it's an integer, then the respective column
         * (1 based index) is returned.
         * @param nameValueEval
         * @param db
         * @return the first column index that matches the given name (or int)
         * @
         */
        private static int GetColumnForTag(ValueEval nameValueEval, TwoDEval db)
        {
            int resultColumn = -1;

            // Numbers as column indicator are allowed, check that.
            if (nameValueEval is NumericValueEval)
            {
                double doubleResultColumn = ((NumericValueEval)nameValueEval).NumberValue;
                resultColumn = (int)doubleResultColumn;
                // Floating comparisions are usually not possible, but should work for 0.0.
                if (doubleResultColumn - resultColumn != 0.0)
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                resultColumn -= 1; // Numbers are 1-based not 0-based.
            }
            else
            {
                resultColumn = GetColumnForName(nameValueEval, db);
            }
            return resultColumn;
        }

        private static int GetColumnForName(ValueEval nameValueEval, TwoDEval db)
        {
            String name = GetStringFromValueEval(nameValueEval);
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
        private static int GetColumnForString(TwoDEval db, String name)
        {
            int resultColumn = -1;
            for (int column = 0; column < db.Width; ++column)
            {
                ValueEval columnNameValueEval = db.GetValue(0, column);
                String columnName = GetStringFromValueEval(columnNameValueEval);
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
        private static bool FullFillsConditions(TwoDEval db, int row, TwoDEval cdb)
        {
            // Only one row must match to accept the input, so rows are ORed.
            // Each row is made up of cells where each cell is a condition,
            // all have to match, so they are ANDed.
            for (int conditionRow = 1; conditionRow < cdb.Height; ++conditionRow)
            {
                bool matches = true;
                for (int column = 0; column < cdb.Width; ++column)
                { // columns are ANDed
                    // Whether the condition column matches a database column, if not it's a
                    // special column that accepts formulas.
                    bool columnCondition = true;
                    ValueEval condition = null;
                    try
                    {
                        // The condition to Apply.
                        condition = solveReference(cdb.GetValue(conditionRow, column));
                    }
                    catch (Exception)
                    {
                        // It might be a special formula, then it is ok if it fails.
                        columnCondition = false;
                    }
                    // If the condition is empty it matches.
                    if (condition is BlankEval)
                        continue;
                    // The column in the DB to apply the condition to.
                    ValueEval targetHeader = solveReference(cdb.GetValue(0, column));
                    targetHeader = solveReference(targetHeader);


                    if (!(targetHeader is StringValueEval))
                        columnCondition = false;
                    else if (GetColumnForName(targetHeader, db) == -1)
                        // No column found, it's again a special column that accepts formulas.
                        columnCondition = false;

                    if (columnCondition == true)
                    { // normal column condition
                        // Should not throw, Checked above.
                        ValueEval target = db.GetValue(
                                row, GetColumnForName(targetHeader, db));
                        // Must be a string.
                        String conditionString = GetStringFromValueEval(condition);
                        if (!testNormalCondition(target, conditionString))
                        {
                            matches = false;
                            break;
                        }
                    }
                    else
                    { // It's a special formula condition.
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
        private static bool testNormalCondition(ValueEval value, String condition)
            {
        if(condition.StartsWith("<")) { // It's a </<= condition.
            String number = condition.Substring(1);
            if(number.StartsWith("=")) {
                number = number.Substring(1);
                return testNumericCondition(value, Operator.smallerEqualThan, number);
            } else {
                return testNumericCondition(value, Operator.smallerThan, number);
            }
        }
        else if(condition.StartsWith(">")) { // It's a >/>= condition.
            String number = condition.Substring(1);
            if(number.StartsWith("=")) {
                number = number.Substring(1);
                return testNumericCondition(value, Operator.largerEqualThan, number);
            } else {
                return testNumericCondition(value, Operator.largerThan, number);
            }
        }
        else if(condition.StartsWith("=")) { // It's a = condition.
            String stringOrNumber = condition.Substring(1);
            // Distinguish between string and number.
            bool itsANumber = false;
            try {
                Int32.Parse(stringOrNumber);
                itsANumber = true;
            } catch (FormatException) { // It's not an int.
                try {
                    Double.Parse(stringOrNumber);
                    itsANumber = true;
                } catch (FormatException) { // It's a string.
                    itsANumber = false;
                }
            }
            if(itsANumber) {
                return testNumericCondition(value, Operator.equal, stringOrNumber);
            } else { // It's a string.
                String valueString = GetStringFromValueEval(value);
                return stringOrNumber.Equals(valueString);
            }
        } else { // It's a text starts-with condition.
            String valueString = GetStringFromValueEval(value);
            return valueString.StartsWith(condition);
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

        /**
         * Takes a ValueEval and tries to retrieve a String value from it.
         * It tries to resolve references if there are any.
         *
         * @param value ValueEval to retrieve the string from.
         * @return String corresponding to the given ValueEval.
         * @If it's not possible to retrieve a String value.
         */
        private static String GetStringFromValueEval(ValueEval value)
        {
            value = solveReference(value);
            if (value is BlankEval)
                return "";
            if (!(value is StringValueEval))
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            return ((StringValueEval)value).StringValue;
        }
    }

}