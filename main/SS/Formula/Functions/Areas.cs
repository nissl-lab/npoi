﻿using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.PTG;
using System;
using System.Collections.Generic;
using System.Text; 
using Cysharp.Text;

namespace NPOI.SS.Formula.Functions
{
    public class Areas : Function
    {
        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length == 0)
            {
                return ErrorEval.VALUE_INVALID;
            }
            try
            {
                ValueEval valueEval = args[0];
                int result = 1;
                if (valueEval is RefListEval refListEval) {
                    result = refListEval.GetList().Count;
                }
                NumberEval numberEval = new NumberEval(new NumberPtg(result));
                return numberEval;
            }
            catch
            {
                return ErrorEval.VALUE_INVALID;
            }
        }
    }
}
