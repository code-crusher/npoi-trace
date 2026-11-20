using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using System;
using NPOI.SS;

namespace NPOI.SS.Formula.Atp
{
    /// <summary>
    /// Implementation of the Excel dynamic array function SEQUENCE,
    /// persisted in files as the future function _xlfn.SEQUENCE.
    /// Generates a list of sequential numbers in an array.
    /// </summary>
    public class SequenceFunction : FreeRefFunction, IArrayFunction
    {
        public static FreeRefFunction instance = new SequenceFunction();

        private SequenceFunction()
        {
            // Enforce singleton
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            return Evaluate(args, ec.RowIndex, ec.ColumnIndex);
        }

        public ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            return Evaluate(args, srcRowIndex, srcColumnIndex);
        }

        private static ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length < 1 || args.Length > 4)
            {
                return ErrorEval.VALUE_INVALID;
            }
            int height = 1;
            int width = 1;
            double start = 1;
            double step = 1;
            try
            {
                switch (args.Length)
                {
                    case 4:
                        step = OperandResolver.CoerceValueToDouble(
                            OperandResolver.GetSingleValue(args[3], srcRowIndex, srcColumnIndex));
                        goto case 3;
                    case 3:
                        start = OperandResolver.CoerceValueToDouble(
                            OperandResolver.GetSingleValue(args[2], srcRowIndex, srcColumnIndex));
                        goto case 2;
                    case 2:
                        width = OperandResolver.CoerceValueToInt(
                            OperandResolver.GetSingleValue(args[1], srcRowIndex, srcColumnIndex));
                        goto case 1;
                    case 1:
                        height = OperandResolver.CoerceValueToInt(
                            OperandResolver.GetSingleValue(args[0], srcRowIndex, srcColumnIndex));
                        break;
                }
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return Evaluate(height, width, start, step, srcRowIndex, srcColumnIndex);
        }

        private static ValueEval Evaluate(int height, int width, double start, double step,
            int srcRowIndex, int srcColumnIndex)
        {
            if (height < 1 || width < 1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            SpreadsheetVersion version = SpreadsheetVersion.EXCEL2007;
            if (height > version.MaxRows - srcRowIndex || width > version.MaxColumns - srcColumnIndex)
            {
                return ErrorEval.NUM_ERROR;
            }
            ValueEval[] values = new ValueEval[height * width];
            for (int i = 0; i < height; i++)
            {
                for (int j = 0; j < width; j++)
                {
                    values[i * width + j] = new NumberEval(start + (i * width + j) * step);
                }
            }
            if (values.Length == 1)
            {
                return values[0];
            }
            return new CacheAreaEval(srcRowIndex, srcColumnIndex,
                srcRowIndex + height - 1, srcColumnIndex + width - 1, values);
        }
    }
}
