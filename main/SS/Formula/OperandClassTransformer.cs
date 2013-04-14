/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula
{

    using System;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;

    /**
     * This class performs 'operand class' transformation. Non-base Tokens are classified into three 
     * operand classes:
     * <ul>
     * <li>reference</li> 
     * <li>value</li> 
     * <li>array</li> 
     * </ul>
     * <p/>
     * 
     * The operand class chosen for each Token depends on the formula type and the Token's place
     * in the formula. If POI Gets the operand class wrong, Excel <em>may</em> interpret the formula
     * incorrectly.  This condition is typically manifested as a formula cell that displays as '#VALUE!',
     * but resolves correctly when the user presses F2, enter.<p/>
     * 
     * The logic implemented here was partially inspired by the description in
     * "OpenOffice.org's Documentation of the Microsoft Excel File Format".  The model presented there
     * seems To be inconsistent with observed Excel behaviour (These differences have not been fully
     * investigated). The implementation in this class Has been heavily modified in order To satisfy
     * concrete examples of how Excel performs the same logic (see TestRVA).<p/>
     * 
     * Hopefully, as Additional important test cases are identified and Added To the test suite, 
     * patterns might become more obvious in this code and allow for simplification.
     * 
     * @author Josh Micich
     */
    class OperandClassTransformer
    {

        private FormulaType _formulaType;

        public OperandClassTransformer(FormulaType formulaType)
        {
            _formulaType = formulaType;
        }

        /**
         * Traverses the supplied formula parse tree, calling <c>Ptg.SetClass()</c> for each non-base
         * Token To Set its operand class.
         */
        public void TransformFormula(ParseNode rootNode)
        {
            byte rootNodeOperandClass;
            switch (_formulaType)
            {
                case FormulaType.Cell:
                    rootNodeOperandClass = Ptg.CLASS_VALUE;
                    break;
                case FormulaType.Array:
                    rootNodeOperandClass = Ptg.CLASS_ARRAY;
                    break;
                case FormulaType.NamedRange:
                case FormulaType.DataValidationList:
                    rootNodeOperandClass = Ptg.CLASS_REF;
                    break;
                default:
                    throw new Exception("Incomplete code - formula type ("
                            + _formulaType + ") not supported yet");

            }
            TransformNode(rootNode, rootNodeOperandClass, false);
        }

        /**
         * @param callerForceArrayFlag <c>true</c> if one of the current node's parents is a 
         * function Ptg which Has been Changed from default 'V' To 'A' type (due To requirements on
         * the function return value).
         */
        private void TransformNode(ParseNode node, byte desiredOperandClass,
                bool callerForceArrayFlag)
        {
            Ptg token = node.GetToken();
            ParseNode[] children = node.GetChildren();
            bool IsSimpleValueFunc = IsSimpleValueFunction(token);

            if (IsSimpleValueFunc)
            {
                bool localForceArray = desiredOperandClass == Ptg.CLASS_ARRAY;
                for (int i = 0; i < children.Length; i++)
                {
                    TransformNode(children[i], desiredOperandClass, localForceArray);
                }
                SetSimpleValueFuncClass((AbstractFunctionPtg)token, desiredOperandClass, callerForceArrayFlag);
                return;
            }
		if (IsSingleArgSum(token)) {
			// Need to process the argument of SUM with transformFunctionNode below
			// so make a dummy FuncVarPtg for that call.
			token = FuncVarPtg.SUM;
			// Note - the tAttrSum token (node.getToken()) is a base
			// token so does not need to have its operand class set
		}
            if (token is ValueOperatorPtg || token is ControlPtg
                || token is MemFuncPtg
				|| token is MemAreaPtg
				|| token is UnionPtg)
            {
                // Value Operator Ptgs and Control are base Tokens, so Token will be unchanged
                // but any child nodes are processed according To desiredOperandClass and callerForceArrayFlag

                // As per OOO documentation Sec 3.2.4 "Token Class Transformation", "Step 1"
                // All direct operands of value operators that are initially 'R' type will 
                // be converted To 'V' type.
                byte localDesiredOperandClass = desiredOperandClass == Ptg.CLASS_REF ? Ptg.CLASS_VALUE : desiredOperandClass;
                for (int i = 0; i < children.Length; i++)
                {
                    TransformNode(children[i], localDesiredOperandClass, callerForceArrayFlag);
                }
                return;
            }
            if (token is AbstractFunctionPtg)
            {
                TransformFunctionNode((AbstractFunctionPtg)token, children, desiredOperandClass, callerForceArrayFlag);
                return;
            }
            if (children.Length > 0)
            {
                //if (token == RangePtg.instance)
                if(token is OperationPtg)
                {
                    // TODO is any Token transformation required under the various ref operators?
                    return;
                }
                throw new InvalidOperationException("Node should not have any children");
            }

            if (token.IsBaseToken)
            {
                // nothing To do
                return;
            }
            token.PtgClass = (TransformClass(token.PtgClass, desiredOperandClass, callerForceArrayFlag));
        }
        private static bool IsSingleArgSum(Ptg token)
        {
            if (token is AttrPtg)
            {
                AttrPtg attrPtg = (AttrPtg)token;
                return attrPtg.IsSum;
            }
            return false;
        }
        private static bool IsSimpleValueFunction(Ptg token)
        {
            if (token is AbstractFunctionPtg)
            {
                AbstractFunctionPtg aptg = (AbstractFunctionPtg)token;
                if (aptg.DefaultOperandClass != Ptg.CLASS_VALUE)
                {
                    return false;
                }
                int numberOfOperands = aptg.NumberOfOperands;
                for (int i = numberOfOperands - 1; i >= 0; i--)
                {
                    if (aptg.GetParameterClass(i) != Ptg.CLASS_VALUE)
                    {
                        return false;
                    }
                }
                return true;
            }
            return false;
        }

        private byte TransformClass(byte currentOperandClass, byte desiredOperandClass,
                bool callerForceArrayFlag)
        {
            switch (desiredOperandClass)
            {
                case Ptg.CLASS_VALUE:
                    if (!callerForceArrayFlag)
                    {
                        return Ptg.CLASS_VALUE;
                    }
                    return Ptg.CLASS_ARRAY;
                    //break;
                // else fall through
                case Ptg.CLASS_ARRAY:
                    return Ptg.CLASS_ARRAY;
                case Ptg.CLASS_REF:
                    if (!callerForceArrayFlag)
                    {
                        return currentOperandClass;
                    }
                    return Ptg.CLASS_REF;
            }
            throw new InvalidOperationException("Unexpected operand class (" + desiredOperandClass + ")");
        }

        private void TransformFunctionNode(AbstractFunctionPtg afp, ParseNode[] children,
                byte desiredOperandClass, bool callerForceArrayFlag)
        {

            bool localForceArrayFlag;
            byte defaultReturnOperandClass = afp.DefaultOperandClass;

            if (callerForceArrayFlag)
            {
                switch (defaultReturnOperandClass)
                {
                    case Ptg.CLASS_REF:
                        if (desiredOperandClass == Ptg.CLASS_REF)
                        {
                            afp.PtgClass = (Ptg.CLASS_REF);
                        }
                        else
                        {
                            afp.PtgClass = (Ptg.CLASS_ARRAY);
                        }
                        localForceArrayFlag = false;
                        break;
                    case Ptg.CLASS_ARRAY:
                        afp.PtgClass = (Ptg.CLASS_ARRAY);
                        localForceArrayFlag = false;
                        break;
                    case Ptg.CLASS_VALUE:
                        afp.PtgClass = (Ptg.CLASS_ARRAY);
                        localForceArrayFlag = true;
                        break;
                    default:
                        throw new InvalidOperationException("Unexpected operand class ("
                                + defaultReturnOperandClass + ")");
                }
            }
            else
            {
                if (defaultReturnOperandClass == desiredOperandClass)
                {
                    localForceArrayFlag = false;
                    // an alternative would have been To for non-base Ptgs To Set their operand class 
                    // from their default, but this would require the call in many subclasses because
                    // the default OC is not known until the end of the constructor
                    afp.PtgClass = (defaultReturnOperandClass);
                }
                else
                {
                    switch (desiredOperandClass)
                    {
                        case Ptg.CLASS_VALUE:
                            // always OK To Set functions To return 'value'
                            afp.PtgClass = (Ptg.CLASS_VALUE);
                            localForceArrayFlag = false;
                            break;
                        case Ptg.CLASS_ARRAY:
                            switch (defaultReturnOperandClass)
                            {
                                case Ptg.CLASS_REF:
                                    afp.PtgClass = (Ptg.CLASS_REF);
                                    //								afp.SetClass(Ptg.CLASS_ARRAY);
                                    break;
                                case Ptg.CLASS_VALUE:
                                    afp.PtgClass = (Ptg.CLASS_ARRAY);
                                    break;
                                default:
                                    throw new InvalidOperationException("Unexpected operand class ("
                                            + defaultReturnOperandClass + ")");
                            }
                            localForceArrayFlag = (defaultReturnOperandClass == Ptg.CLASS_VALUE);
                            break;
                        case Ptg.CLASS_REF:
                            switch (defaultReturnOperandClass)
                            {
                                case Ptg.CLASS_ARRAY:
                                    afp.PtgClass=(Ptg.CLASS_ARRAY);
                                    break;
                                case Ptg.CLASS_VALUE:
                                    afp.PtgClass=(Ptg.CLASS_VALUE);
                                    break;
                                default:
                                    throw new InvalidOperationException("Unexpected operand class ("
                                            + defaultReturnOperandClass + ")");
                            }
                            localForceArrayFlag = false;
                            break;
                        default:
                            throw new InvalidOperationException("Unexpected operand class ("
                                    + desiredOperandClass + ")");
                    }

                }
            }

            for (int i = 0; i < children.Length; i++)
            {
                ParseNode child = children[i];
                byte paramOperandClass = afp.GetParameterClass(i);
                TransformNode(child, paramOperandClass, localForceArrayFlag);
            }
        }

        private void SetSimpleValueFuncClass(AbstractFunctionPtg afp,
                byte desiredOperandClass, bool callerForceArrayFlag)
        {

            if (callerForceArrayFlag || desiredOperandClass == Ptg.CLASS_ARRAY)
            {
                afp.PtgClass = (Ptg.CLASS_ARRAY);
            }
            else
            {
                afp.PtgClass = (Ptg.CLASS_VALUE);
            }
        }
    }
}
