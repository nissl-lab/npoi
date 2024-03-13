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
using NPOI.SS.UserModel;
using System;
using System.Text;
using NPOI.Util;
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Text.RegularExpressions;

namespace NPOI.XSSF.UserModel
{

    /**
     * @author <a href="rjankiraman@emptoris.com">Radhakrishnan J</a>
     *
     */
    public class XSSFDataValidationConstraint : IDataValidationConstraint
    {
        /**
         * Excel validation constraints with static lists are delimited with optional whitespace and the Windows List Separator,
         * which is typically comma, but can be changed by users.  POI will just assume comma.
         */
        private static String LIST_SEPARATOR = ",";
        private static Regex LIST_SPLIT_REGEX = new Regex("\\s*" + LIST_SEPARATOR + "\\s*");
        private static String QUOTE = "\"";
        private String formula1;
        private String formula2;
        private int validationType = -1;
        private int operator1 = -1;
        private String[] explicitListOfValues;

        public XSSFDataValidationConstraint(String[] explicitListOfValues)
        {
            if (explicitListOfValues == null || explicitListOfValues.Length == 0)
            {
                throw new ArgumentException("List validation with explicit values must specify at least one value");
            }
            this.validationType = ValidationType.LIST;
            ExplicitListValues = (explicitListOfValues);

            Validate();
        }

        public XSSFDataValidationConstraint(int validationType, String formula1)
            : base()
        {

            Formula1 = (formula1);
            this.validationType = validationType;
            Validate();
        }



        public XSSFDataValidationConstraint(int validationType, int operator1, String formula1)
            : base()
        {

            Formula1 = (formula1);
            this.validationType = validationType;
            this.operator1 = operator1;
            Validate();
        }

        /// <summary>
        /// This is the constructor called using the OOXML raw data.  Excel overloads formula1 to also encode explicit value lists,
        /// so this constructor has to check for and parse that syntax.
        /// </summary>
        /// <param name="validationType"></param>
        /// <param name="operator1"></param>
        /// <param name="formula1">Overloaded: formula1 or list of explicit values</param>
        /// <param name="formula2">formula1 is a list of explicit values, this is ignored: use <code>null</code></param>
        public XSSFDataValidationConstraint(int validationType, int operator1, String formula1, String formula2)
            : base()
        {
            //removes leading equals sign if present
            Formula1 = (formula1);
            Formula2 = (formula2);
            this.validationType = validationType;
            this.operator1 = operator1;

            Validate();

            //FIXME: Need to confirm if this is not a formula.
            // empirical testing shows Excel saves explicit lists surrounded by double quotes, 
            // range formula expressions can't start with quotes (I think - anyone have a creative counter example?)
            if (ValidationType.LIST == validationType
                    && this.formula1 != null
                    && IsQuoted(this.formula1))
            {
                explicitListOfValues = LIST_SPLIT_REGEX.Split(Unquote(this.formula1));
            }
        }

        public String[] ExplicitListValues
        {
            get
            {
                return explicitListOfValues;
            }
            set 
            {
                this.explicitListOfValues = value;
                // for OOXML we need to set formula1 to the quoted csv list of values (doesn't appear documented, but that's where Excel puts its lists)
                // further, Excel has no escaping for commas in explicit lists, so we don't need to worry about that.
                if (explicitListOfValues != null && explicitListOfValues.Length > 0)
                {
                    StringBuilder builder = new StringBuilder(QUOTE);
                    for (int i = 0; i < value.Length; i++)
                    {
                        String string1 = value[i];
                        if (builder.Length > 1)
                        {
                            builder.Append(LIST_SEPARATOR);
                        }
                        builder.Append(string1);
                    }
                    builder.Append(QUOTE);
                    Formula1 = builder.ToString();
                }
            }
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidationConstraint#getFormula1()
         */
        public String Formula1
        {
            get
            {
                return formula1;
            }
            set 
            {
                this.formula1 = RemoveLeadingEquals(value);
            }
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidationConstraint#getFormula2()
         */
        public String Formula2
        {
            get
            {
                return formula2;
            }
            set 
            {
                this.formula2 = RemoveLeadingEquals(value);
            }
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidationConstraint#getOperator()
         */
        public int Operator
        {
            get
            {
                return operator1;
            }
            set 
            {
                this.operator1 = value;
            }
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidationConstraint#getValidationType()
         */
        public int GetValidationType()
        {
            return validationType;
        }

        protected static string RemoveLeadingEquals(string formula1)
        {
            return IsFormulaEmpty(formula1) ? formula1 : formula1[0] == '=' ? formula1.Substring(1) : formula1;
        }

        private static bool IsQuoted(string s)
        {
            return s.StartsWith(QUOTE) && s.EndsWith(QUOTE);
        }
        private static String Unquote(string s)
        {
            // removes leading and trailing quotes from a quoted string
            if (IsQuoted(s))
            {
                return s.Substring(1, s.Length - 2);
            }
            return s;
        }
        protected static bool IsFormulaEmpty(string formula1)
        {
            return formula1 == null || formula1.Trim().Length == 0;
        }

        public void Validate()
        {
            if (validationType == ValidationType.ANY)
            {
                return;
            }

            if (validationType == ValidationType.LIST)
            {
                if (IsFormulaEmpty(formula1))
                {
                    throw new ArgumentException("A valid formula or a list of values must be specified for list validation.");
                }
            }
            else
            {
                if (IsFormulaEmpty(formula1))
                {
                    throw new ArgumentException("Formula is not specified. Formula is required for all validation types except explicit list validation.");
                }

                if (validationType != ValidationType.FORMULA)
                {
                    if (operator1 == -1)
                    {
                        throw new ArgumentException("This validation type requires an operator to be specified.");
                    }
                    else if ((operator1 == OperatorType.BETWEEN || operator1 == OperatorType.NOT_BETWEEN) && IsFormulaEmpty(formula2))
                    {
                        throw new ArgumentException("Between and not between comparisons require two formulae to be specified.");
                    }
                }
            }
        }


        public string PrettyPrint()
        {
            StringBuilder builder = new StringBuilder();
            ST_DataValidationType vt = XSSFDataValidation.validationTypeMappings[validationType];
            Enum ot = XSSFDataValidation.operatorTypeMappings[operator1];
            builder.Append(vt);
            builder.Append(' ');
            if (validationType != ValidationType.ANY)
            {
                if (validationType != ValidationType.LIST
                    && validationType != ValidationType.FORMULA)
                {
                    builder.Append(LIST_SEPARATOR).Append(ot).Append(", ");
                }
                if (validationType == ValidationType.LIST && explicitListOfValues != null)
                {
                    builder.Append(QUOTE).Append(Arrays.AsList(explicitListOfValues)).Append(QUOTE).Append(' ');
                }
                else
                {
                    builder.Append(QUOTE).Append(formula1).Append(QUOTE).Append(' ');
                }
                if (formula2 != null)
                {
                    builder.Append(QUOTE).Append(formula2).Append(QUOTE).Append(' ');
                }
            }
            return builder.ToString();
        }
    }
}

