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




using NPOI.ss.usermodel.DataValidation;
using NPOI.ss.usermodel.DataValidationConstraint;
using NPOI.ss.usermodel.DataValidationHelper;
using NPOI.ss.usermodel.DataValidationConstraint.ValidationType;
using NPOI.ss.util.CellRangeAddress;
using NPOI.ss.util.CellRangeAddressList;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationType;

/**
 * @author <a href="rjankiraman@emptoris.com">Radhakrishnan J</a>
 *
 */
public class XSSFDataValidationHelper : DataValidationHelper {
	private XSSFSheet xssfSheet;
	
    
    public XSSFDataValidationHelper(XSSFSheet xssfSheet) {
		base();
		this.xssfSheet = xssfSheet;
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationHelper#CreateDateConstraint(int, java.lang.String, java.lang.String, java.lang.String)
	 */
	public DataValidationConstraint CreateDateConstraint(int operatorType, String formula1, String formula2, String dateFormat) {
		return new XSSFDataValidationConstraint(ValidationType.DATE, operatorType,formula1, formula2);
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationHelper#CreateDecimalConstraint(int, java.lang.String, java.lang.String)
	 */
	public DataValidationConstraint CreateDecimalConstraint(int operatorType, String formula1, String formula2) {
		return new XSSFDataValidationConstraint(ValidationType.DECIMAL, operatorType,formula1, formula2);
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationHelper#CreateExplicitListConstraint(java.lang.String[])
	 */
	public DataValidationConstraint CreateExplicitListConstraint(String[] listOfValues) {
		return new XSSFDataValidationConstraint(listOfValues);
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationHelper#CreateFormulaListConstraint(java.lang.String)
	 */
	public DataValidationConstraint CreateFormulaListConstraint(String listFormula) {
		return new XSSFDataValidationConstraint(ValidationType.LIST, listFormula);
	}

	
	
	public DataValidationConstraint CreateNumericConstraint(int validationType, int operatorType, String formula1, String formula2) {
		if( validationType==ValidationType.INTEGER) {
			return CreateintConstraint(operatorType, formula1, formula2);
		} else if ( validationType==ValidationType.DECIMAL) {
			return CreateDecimalConstraint(operatorType, formula1, formula2);
		} else if ( validationType==ValidationType.TEXT_LENGTH) {
			return CreateTextLengthConstraint(operatorType, formula1, formula2);
		}
		return null;
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationHelper#CreateintConstraint(int, java.lang.String, java.lang.String)
	 */
	public DataValidationConstraint CreateintConstraint(int operatorType, String formula1, String formula2) {
		return new XSSFDataValidationConstraint(ValidationType.INTEGER, operatorType,formula1,formula2);
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationHelper#CreateTextLengthConstraint(int, java.lang.String, java.lang.String)
	 */
	public DataValidationConstraint CreateTextLengthConstraint(int operatorType, String formula1, String formula2) {
		return new XSSFDataValidationConstraint(ValidationType.TEXT_LENGTH, operatorType,formula1,formula2);
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationHelper#CreateTimeConstraint(int, java.lang.String, java.lang.String, java.lang.String)
	 */
	public DataValidationConstraint CreateTimeConstraint(int operatorType, String formula1, String formula2) {
		return new XSSFDataValidationConstraint(ValidationType.TIME, operatorType,formula1,formula2);
	}

	public DataValidationConstraint CreateCustomConstraint(String formula) {
		return new XSSFDataValidationConstraint(ValidationType.FORMULA, formula);
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidationHelper#CreateValidation(NPOI.ss.usermodel.DataValidationConstraint, NPOI.ss.util.CellRangeAddressList)
	 */
	public DataValidation CreateValidation(DataValidationConstraint constraint, CellRangeAddressList cellRangeAddressList) {
		XSSFDataValidationConstraint dataValidationConstraint = (XSSFDataValidationConstraint)constraint;
		CTDataValidation newDataValidation = CTDataValidation.Factory.newInstance();

		int validationType = constraint.GetValidationType();
		switch(validationType) {
			case DataValidationConstraint.ValidationType.LIST:
		    	newDataValidation.SetType(STDataValidationType.LIST);
				newDataValidation.SetFormula1(constraint.GetFormula1());				
		    	break;
			case DataValidationConstraint.ValidationType.ANY:				
				newDataValidation.SetType(STDataValidationType.NONE);				
				break;
			case DataValidationConstraint.ValidationType.TEXT_LENGTH:
				newDataValidation.SetType(STDataValidationType.TEXT_LENGTH);
				break;				
			case DataValidationConstraint.ValidationType.DATE:
				newDataValidation.SetType(STDataValidationType.DATE);
				break;				
			case DataValidationConstraint.ValidationType.INTEGER:
				newDataValidation.SetType(STDataValidationType.WHOLE);
				break;				
			case DataValidationConstraint.ValidationType.DECIMAL:
				newDataValidation.SetType(STDataValidationType.DECIMAL);
				break;				
			case DataValidationConstraint.ValidationType.TIME:
				newDataValidation.SetType(STDataValidationType.TIME);
				break;
			case DataValidationConstraint.ValidationType.FORMULA:
				newDataValidation.SetType(STDataValidationType.CUSTOM);
				break;
			default:
				newDataValidation.SetType(STDataValidationType.NONE);				
		}
		
		if (validationType!=ValidationType.ANY && validationType!=ValidationType.LIST) {
			newDataValidation.SetOperator(XSSFDataValidation.operatorTypeMappings.Get(constraint.GetOperator()));			
			if (constraint.GetFormula1() != null) {
				newDataValidation.SetFormula1(constraint.GetFormula1());
			}
			if (constraint.GetFormula2() != null) {
				newDataValidation.SetFormula2(constraint.GetFormula2());
			}
		}
		
		CellRangeAddress[] cellRangeAddresses = cellRangeAddressList.GetCellRangeAddresses();
		List<String> sqref = new List<String>();
		for (int i = 0; i < cellRangeAddresses.Length; i++) {
			CellRangeAddress cellRangeAddress = cellRangeAddresses[i];
			sqref.Add(cellRangeAddress.formatAsString());
		}
		newDataValidation.SetSqref(sqref);
		
		return new XSSFDataValidation(dataValidationConstraint,cellRangeAddressList,newDataValidation);
	}
}


