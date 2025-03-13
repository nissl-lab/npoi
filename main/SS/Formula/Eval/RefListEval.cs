using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Eval
{
    public class RefListEval:ValueEval
    {
        private  List<ValueEval> list = new List<ValueEval>();

        public RefListEval(ValueEval v1, ValueEval v2)
        {
            Add(v1);
            Add(v2);
        }

        private void Add(ValueEval v)
        {
            // flatten multiple nested RefListEval
            if (v is RefListEval eval) {
                list.AddRange(eval.list);
            } else
            {
                list.Add(v);
            }
        }

        public List<ValueEval> GetList()
        {
            return list;
        }
    }
}
