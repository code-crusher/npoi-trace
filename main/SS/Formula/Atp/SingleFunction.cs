using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using System;

namespace NPOI.SS.Formula.Atp
{
    /// <summary>
    /// Implementation of the Excel implicit intersection operator '@',
    /// persisted in files as the future function _xlfn.SINGLE.
    /// Returns a single value from the referenced area, selected by
    /// implicit intersection with the coordinates of the calling cell.
    /// </summary>
    public class SingleFunction : FreeRefFunction, IArrayFunction
    {
        public static FreeRefFunction instance = new SingleFunction();

        private SingleFunction()
        {
            // Enforce singleton
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            try
            {
                return OperandResolver.GetSingleValue(args[0], ec.RowIndex, ec.ColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        public ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length != 1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            try
            {
                return OperandResolver.GetSingleValue(args[0], srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
    }
}
