using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using System;
using NPOI.SS;

namespace NPOI.SS.Formula.Atp
{
    /// <summary>
    /// Implementation of the Excel dynamic array function RANDARRAY,
    /// persisted in files as the future function _xlfn.RANDARRAY.
    /// Returns an array of random numbers.
    /// </summary>
    public class RandArrayFunction : FreeRefFunction, IArrayFunction
    {
        public static FreeRefFunction instance = new RandArrayFunction();

        private readonly Random _rnd;

        private RandArrayFunction()
        {
            // Enforce singleton
            _rnd = new Random();
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            return EvaluateArgs(args, ec.RowIndex, ec.ColumnIndex);
        }

        public ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            return EvaluateArgs(args, srcRowIndex, srcColumnIndex);
        }

        private ValueEval EvaluateArgs(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length > 5)
            {
                return ErrorEval.VALUE_INVALID;
            }
            int height = 1;
            int width = 1;
            double min = 0;
            double max = 1;
            bool wholeNumber = false;
            try
            {
                switch (args.Length)
                {
                    case 5:
                        wholeNumber = OperandResolver.CoerceValueToBoolean(
                            OperandResolver.GetSingleValue(args[4], srcRowIndex, srcColumnIndex), false) == true;
                        goto case 4;
                    case 4:
                        max = OperandResolver.CoerceValueToDouble(
                            OperandResolver.GetSingleValue(args[3], srcRowIndex, srcColumnIndex));
                        goto case 3;
                    case 3:
                        min = OperandResolver.CoerceValueToDouble(
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
                    case 0:
                        break;
                }
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return Evaluate(height, width, min, max, wholeNumber, srcRowIndex, srcColumnIndex);
        }

        private ValueEval Evaluate(int height, int width, double min, double max,
            bool wholeNumber, int srcRowIndex, int srcColumnIndex)
        {
            if (height < 1 || width < 1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            if (max < min)
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
                    double value = min + _rnd.NextDouble() * (max - min);
                    values[i * width + j] = new NumberEval(wholeNumber ? Math.Floor(value) : value);
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
