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
using NPOI.ss.usermodel.DataValidationConstraint.ValidationType;
using NPOI.ss.util.CellRangeAddress;
using NPOI.ss.util.CellRangeAddressList;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationErrorStyle;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationType;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.Enum;

/**
 * @author <a href="rjankiraman@emptoris.com">Radhakrishnan J</a>
 *
 */
public class XSSFDataValidation : DataValidation {
	private CTDataValidation ctDdataValidation;
	private XSSFDataValidationConstraint validationConstraint;
	private CellRangeAddressList regions;

    static Dictionary<int,STDataValidationOperator.Enum> operatorTypeMappings = new Dictionary<int,STDataValidationOperator.Enum>();
	static Dictionary<STDataValidationOperator.Enum,int> operatorTypeReverseMappings = new Dictionary<STDataValidationOperator.Enum,int>();
	static Dictionary<int,STDataValidationType.Enum> validationTypeMappings = new Dictionary<int,STDataValidationType.Enum>();
	static Dictionary<STDataValidationType.Enum,int> validationTypeReverseMappings = new Dictionary<STDataValidationType.Enum,int>();
    static Dictionary<int,STDataValidationErrorStyle.Enum> errorStyleMappings = new Dictionary<int,STDataValidationErrorStyle.Enum>();
    static {
		errorStyleMappings.Put(DataValidation.ErrorStyle.INFO, STDataValidationErrorStyle.INFORMATION);
		errorStyleMappings.Put(DataValidation.ErrorStyle.STOP, STDataValidationErrorStyle.STOP);
		errorStyleMappings.Put(DataValidation.ErrorStyle.WARNING, STDataValidationErrorStyle.WARNING);
    }
	
    
	static {
		operatorTypeMappings.Put(DataValidationConstraint.OperatorType.BETWEEN,STDataValidationOperator.BETWEEN);
		operatorTypeMappings.Put(DataValidationConstraint.OperatorType.NOT_BETWEEN,STDataValidationOperator.NOT_BETWEEN);
		operatorTypeMappings.Put(DataValidationConstraint.OperatorType.EQUAL,STDataValidationOperator.EQUAL);
		operatorTypeMappings.Put(DataValidationConstraint.OperatorType.NOT_EQUAL,STDataValidationOperator.NOT_EQUAL);
		operatorTypeMappings.Put(DataValidationConstraint.OperatorType.GREATER_THAN,STDataValidationOperator.GREATER_THAN);    	
		operatorTypeMappings.Put(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL,STDataValidationOperator.GREATER_THAN_OR_EQUAL);
		operatorTypeMappings.Put(DataValidationConstraint.OperatorType.LESS_THAN,STDataValidationOperator.LESS_THAN);    	
		operatorTypeMappings.Put(DataValidationConstraint.OperatorType.LESS_OR_EQUAL,STDataValidationOperator.LESS_THAN_OR_EQUAL);
		
		foreach( Map.Entry<int,STDataValidationOperator.Enum> entry in operatorTypeMappings.entrySet() ) {
			operatorTypeReverseMappings.Put(entry.GetValue(),entry.GetKey());
		}
	}

	static {
		validationTypeMappings.Put(DataValidationConstraint.ValidationType.FORMULA,STDataValidationType.CUSTOM);
		validationTypeMappings.Put(DataValidationConstraint.ValidationType.DATE,STDataValidationType.DATE);
		validationTypeMappings.Put(DataValidationConstraint.ValidationType.DECIMAL,STDataValidationType.DECIMAL);    	
		validationTypeMappings.Put(DataValidationConstraint.ValidationType.LIST,STDataValidationType.LIST); 
		validationTypeMappings.Put(DataValidationConstraint.ValidationType.ANY,STDataValidationType.NONE);
		validationTypeMappings.Put(DataValidationConstraint.ValidationType.TEXT_LENGTH,STDataValidationType.TEXT_LENGTH);
		validationTypeMappings.Put(DataValidationConstraint.ValidationType.TIME,STDataValidationType.TIME);  
		validationTypeMappings.Put(DataValidationConstraint.ValidationType.INTEGER,STDataValidationType.WHOLE);
		
		foreach( Map.Entry<int,STDataValidationType.Enum> entry in validationTypeMappings.entrySet() ) {
			validationTypeReverseMappings.Put(entry.GetValue(),entry.GetKey());
		}
	}

	
	XSSFDataValidation(CellRangeAddressList regions,CTDataValidation ctDataValidation) {
		base();
		this.validationConstraint = GetConstraint(ctDataValidation);
		this.ctDdataValidation = ctDataValidation;
		this.regions = regions;
		this.ctDdataValidation.SetErrorStyle(STDataValidationErrorStyle.STOP);
		this.ctDdataValidation.SetAllowBlank(true);
	}	

	public XSSFDataValidation(XSSFDataValidationConstraint constraint,CellRangeAddressList regions,CTDataValidation ctDataValidation) {
		base();
		this.validationConstraint = constraint;
		this.ctDdataValidation = ctDataValidation;
		this.regions = regions;
		this.ctDdataValidation.SetErrorStyle(STDataValidationErrorStyle.STOP);
		this.ctDdataValidation.SetAllowBlank(true);
	}
 
	CTDataValidation GetCtDdataValidation() {
		return ctDdataValidation;
	}



	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#CreateErrorBox(java.lang.String, java.lang.String)
	 */
	public void CreateErrorBox(String title, String text) {
		ctDdataValidation.SetErrorTitle(title);
		ctDdataValidation.SetError(text);
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#CreatePromptBox(java.lang.String, java.lang.String)
	 */
	public void CreatePromptBox(String title, String text) {
		ctDdataValidation.SetPromptTitle(title);
		ctDdataValidation.SetPrompt(text);
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#getEmptyCellAllowed()
	 */
	public bool GetEmptyCellAllowed() {
		return ctDdataValidation.GetAllowBlank();
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#getErrorBoxText()
	 */
	public String GetErrorBoxText() {
		return ctDdataValidation.GetError();
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#getErrorBoxTitle()
	 */
	public String GetErrorBoxTitle() {
		return ctDdataValidation.GetErrorTitle();
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#getErrorStyle()
	 */
	public int GetErrorStyle() {
		return ctDdataValidation.GetErrorStyle().intValue();
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#getPromptBoxText()
	 */
	public String GetPromptBoxText() {
		return ctDdataValidation.GetPrompt();
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#getPromptBoxTitle()
	 */
	public String GetPromptBoxTitle() {
		return ctDdataValidation.GetPromptTitle();
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#getShowErrorBox()
	 */
	public bool GetShowErrorBox() {
		return ctDdataValidation.GetShowErrorMessage();
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#getShowPromptBox()
	 */
	public bool GetShowPromptBox() {
		return ctDdataValidation.GetShowInputMessage();
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#getSuppressDropDownArrow()
	 */
	public bool GetSuppressDropDownArrow() {
		return !ctDdataValidation.GetShowDropDown();
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#getValidationConstraint()
	 */
	public DataValidationConstraint GetValidationConstraint() {
		return validationConstraint;
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#setEmptyCellAllowed(bool)
	 */
	public void SetEmptyCellAllowed(bool allowed) {
		ctDdataValidation.SetAllowBlank(allowed);
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#setErrorStyle(int)
	 */
	public void SetErrorStyle(int errorStyle) {
		ctDdataValidation.SetErrorStyle(errorStyleMappings.Get(errorStyle));
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#setShowErrorBox(bool)
	 */
	public void SetShowErrorBox(bool Show) {
		ctDdataValidation.SetShowErrorMessage(Show);
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#setShowPromptBox(bool)
	 */
	public void SetShowPromptBox(bool Show) {
		ctDdataValidation.SetShowInputMessage(Show);
	}

	/* (non-Javadoc)
	 * @see NPOI.ss.usermodel.DataValidation#setSuppressDropDownArrow(bool)
	 */
	public void SetSuppressDropDownArrow(bool suppress) {
		if (validationConstraint.GetValidationType()==ValidationType.LIST) {
			ctDdataValidation.SetShowDropDown(!suppress);
		}
	}

	public CellRangeAddressList GetRegions() {
		return regions;
	}
	
	public String prettyPrint() {
		StringBuilder builder = new StringBuilder();
		foreach(CellRangeAddress Address in regions.GetCellRangeAddresses()) {
			builder.Append(Address.formatAsString());
		}
		builder.Append(" => ");
		builder.Append(this.validationConstraint.prettyPrint());	
		return builder.ToString();
	}
	
    private XSSFDataValidationConstraint GetConstraint(CTDataValidation ctDataValidation) {
    	XSSFDataValidationConstraint constraint = null;
    	String formula1 = ctDataValidation.GetFormula1();
    	String formula2 = ctDataValidation.GetFormula2();
    	Enum operator = ctDataValidation.GetOperator();
    	org.Openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationType.Enum type = ctDataValidation.GetType();
		int validationType = XSSFDataValidation.validationTypeReverseMappings.Get(type);
		int operatorType = XSSFDataValidation.operatorTypeReverseMappings.Get(operator);
		constraint = new XSSFDataValidationConstraint(validationType,operatorType, formula1,formula2);
    	return constraint;
    }
}


