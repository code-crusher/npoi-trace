using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;

namespace NPOI.SS.Formula.Atp
{
    /// <summary>
    /// Implementation of the Excel spilled range operator '#',
    /// persisted in files as the future function _xlfn.ANCHORARRAY.
    /// Resolves the range spilled by the dynamic array formula of the
    /// referenced anchor cell by evaluating the anchor formula in array mode.
    /// </summary>
    public class AnchorArrayFunction : FreeRefFunction
    {
        public static FreeRefFunction instance = new AnchorArrayFunction();

        private AnchorArrayFunction()
        {
            // Enforce singleton
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 1 || args[0] is not RefEval anchorRef)
            {
                return ErrorEval.VALUE_INVALID;
            }
            IEvaluationSheet evalSheet = ec.GetWorkbook().GetSheet(anchorRef.FirstSheetIndex);
            IEvaluationCell anchorCell = evalSheet.GetCell(anchorRef.Row, anchorRef.Column);
            if (anchorCell == null || anchorCell.CellType != CellType.Formula)
            {
                // no dynamic array is anchored at the referenced cell
                return ErrorEval.REF_INVALID;
            }
            return ec.BookEvaluator.EvaluateAnchorArray(anchorCell, ec.Tracker);
        }
    }
}
