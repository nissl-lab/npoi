using System;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;

namespace NPOI.SS.Formula
{
    public class UserDefinedFunction : FreeRefFunction
    {

        public static FreeRefFunction instance = new UserDefinedFunction();

        private UserDefinedFunction()
        {
            // enforce Singleton
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            int nIncomingArgs = args.Length;
            if (nIncomingArgs < 1)
            {
                throw new Exception("function name argument missing");
            }

            ValueEval nameArg = args[0];
            String functionName = string.Empty ;
            if (nameArg is FunctionNameEval)
            {
                functionName = ((FunctionNameEval)nameArg).FunctionName;
            }
            else
            {
                throw new Exception("First argument should be a NameEval, but got ("
                        + nameArg.GetType().Name + ")");
            }
            FreeRefFunction targetFunc = ec.FindUserDefinedFunction(functionName);
            if (targetFunc == null)
            {
                throw new NotImplementedFunctionException(functionName);
            }
            int nOutGoingArgs = nIncomingArgs - 1;
            ValueEval[] outGoingArgs = new ValueEval[nOutGoingArgs];
            Array.Copy(args, 1, outGoingArgs, 0, nOutGoingArgs);
            return targetFunc.Evaluate(outGoingArgs, ec);
        }
    }
}
