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
namespace NPOI.xssf.usermodel;



using NPOI.ss.usermodel.DataValidationConstraint;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationType;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.Enum;

/**
 * @author <a href="rjankiraman@emptoris.com">Radhakrishnan J</a>
 *
 */
public class XSSFDataValidationConstraint : DataValidationConstraint {
	private String formula1;
	private String formula2;
	private int validationType = -1;
	private int operator = -1;
	private String[] explicitListOfValues;

	public XSSFDataValidationConstraint(String[] explicitListOfValues) {
		if( explicitListOfValues==null || explicitListOfValues.Length==0) {
			throw new ArgumentException("List validation with explicit values must specify at least one value");
		}
		this.validationType = ValidationType.LIST;
	 SetExplicitListValues(explicitListOfValues);
		
		validate();
	}
	
	public XSSFDataValidationConstraint(int validationType,String formula1) {
		base();
	 SetFormula1(formula1);
		this.validationType = validationType;
		validate();
	}



	public XSSFDataValidationConstraint(int validationType, int operator,String formula1) {
		base();
	 SetFormula1(formula1);
		this.validationType = validationType;
		this.operator = operator;
		validate();
	}

	public XSSFDataValidationConstraint(int validationType, int operator,String formula1, String formula2) {
		base();
	 SetFormula1(formula1);
	 SetFormula2(formula2);
		this.validationType = validationType;
		this.operator = operator;
		
		validate();
		
		//FIXME: Need to confirm if this is not a formula.
		if( ValidationType.LIST==validationType) {
			explicitListOfValues = formula1.split(",");
		}
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationConstraint#getExplicitListValues()
	 */
	public String[] GetExplicitListValues() {
		return explicitListOfValues;
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationConstraint#getFormula1()
	 */
	public String GetFormula1() {
		return formula1;
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationConstraint#getFormula2()
	 */
	public String GetFormula2() {
		return formula2;
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationConstraint#getOperator()
	 */
	public int GetOperator() {
		return operator;
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationConstraint#getValidationType()
	 */
	public int GetValidationType() {
		return validationType;
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationConstraint#setExplicitListValues(java.lang.String[])
	 */
	public void SetExplicitListValues(String[] explicitListValues) {
		this.explicitListOfValues = explicitListValues;
		if( explicitListOfValues!=null && explicitListOfValues.Length > 0 ) {
			StringBuilder builder = new StringBuilder("\"");
			for (int i = 0; i < explicitListValues.Length; i++) {
				String string = explicitListValues[i];
				if( builder.Length > 1) {
					builder.Append(",");
				}
				builder.Append(string);
			}
			builder.Append("\"");
		 SetFormula1(builder.ToString());			
		}
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationConstraint#setFormula1(java.lang.String)
	 */
	public void SetFormula1(String formula1) {
		this.formula1 = RemoveLeadingEquals(formula1);
	}

	protected String RemoveLeadingEquals(String formula1) {
		return isFormulaEmpty(formula1) ? formula1 : formula1[0]=='=' ? formula1.Substring(1) : formula1;
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationConstraint#setFormula2(java.lang.String)
	 */
	public void SetFormula2(String formula2) {
		this.formula2 = RemoveLeadingEquals(formula2);
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationConstraint#setOperator(int)
	 */
	public void SetOperator(int operator) {
		this.operator = operator;
	}

	public void validate() {
		if (validationType==ValidationType.ANY) {
			return;
		}
		
		if (validationType==ValidationType.LIST ) {
			if (isFormulaEmpty(formula1)) {
				throw new ArgumentException("A valid formula or a list of values must be specified for list validation.");
			}
		} else  {
			if( isFormulaEmpty(formula1) ) {
				throw new ArgumentException("Formula is not specified. Formula is required for all validation types except explicit list validation.");
			}
			
			if( validationType!= ValidationType.FORMULA ) {
				if (operator==-1) {
					throw new ArgumentException("This validation type requires an operator to be specified.");
				} else if (( operator==OperatorType.BETWEEN || operator==OperatorType.NOT_BETWEEN) && isFormulaEmpty(formula2)) {
					throw new ArgumentException("Between and not between comparisons require two formulae to be specified.");
				}
			}
		}
	}

	protected bool IsFormulaEmpty(String formula1) {
		return formula1 == null || formula1.Trim().Length==0;
	}
	
	public String prettyPrint() {
		StringBuilder builder = new StringBuilder();
		STDataValidationType.Enum vt = XSSFDataValidation.validationTypeMappings.Get(validationType);
		Enum ot = XSSFDataValidation.operatorTypeMappings.Get(operator);
		builder.Append(vt);
		builder.Append(' ');
		if (validationType!=ValidationType.ANY) {
			if (validationType != ValidationType.LIST && validationType != ValidationType.ANY && validationType != ValidationType.FORMULA) {
				builder.Append(",").Append(ot).append(", ");
			}
		 String QUOTE = "";
			if (validationType == ValidationType.LIST && explicitListOfValues != null) {
				builder.Append(QUOTE).Append(Arrays.asList(explicitListOfValues)).append(QUOTE).append(' ');
			} else {
				builder.Append(QUOTE).Append(formula1).append(QUOTE).append(' ');
			}
			if (formula2 != null) {
				builder.Append(QUOTE).Append(formula2).append(QUOTE).append(' ');
			}
		}
		return builder.ToString();
	}
}


