/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/
namespace NPOI.SS.Formula.Functions
{
    /**
     * Implementation for the PMT() Excel function.<p/>
     * 
     * <b>Syntax:</b><br/>
     * <b>PMT</b>(<b>rate</b>, <b>nper</b>, <b>pv</b>, fv, type)<p/>
     * 
     * Returns the constant repayment amount required for a loan assuming a constant interest rate.<p/>
     * 
     * <b>rate</b> the loan interest rate.<br/>
     * <b>nper</b> the number of loan repayments.<br/>
     * <b>pv</b> the present value of the future payments (or principle).<br/>
     * <b>fv</b> the future value (default zero) surplus cash at the end of the loan lifetime.<br/>
     * <b>type</b> whether payments are due at the beginning(1) or end(0 - default) of each payment period.<br/>
     * 
     */
    public class Pmt : FinanceFunction
    {
        public override double Evaluate(double rate, double arg1, double arg2, double arg3, bool type)
        {
            return FinanceLib.pmt(rate, arg1, arg2, arg3, type);
        }
    }
}