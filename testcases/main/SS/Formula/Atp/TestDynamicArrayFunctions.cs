/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace TestCases.SS.Formula.Atp
{
    using System;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /**
     * Tests for dynamic array functions and the implicit intersection ('@')
     * and spilled range ('#') operators, see https://github.com/nissl-lab/npoi/issues/663.
     *
     * Excel persists '@X' as the future function _xlfn.SINGLE(X) and 'X#' as
     * _xlfn.ANCHORARRAY(X).
     */
    [TestFixture]
    public class TestDynamicArrayFunctions
    {
        private HSSFWorkbook _wb;
        private ISheet _sheet;
        private HSSFFormulaEvaluator _fe;
        private ICell _testCell;

        [SetUp]
        public void SetUp()
        {
            _wb = new HSSFWorkbook();
            _sheet = _wb.CreateSheet();
            _fe = new HSSFFormulaEvaluator(_wb);
            // a plain cell for entering the formula under test, far away from the test data
            _testCell = _sheet.CreateRow(30).CreateCell(10);
        }

        [TearDown]
        public void TearDown()
        {
            _wb.Close();
        }

        private void AssertNumber(string formula, double expected, double delta = 0.0)
        {
            TestCases.SS.Util.Utils.AssertDouble(_fe, _testCell, formula, expected, delta);
        }

        private void AssertError(string formula, FormulaError expectedError)
        {
            TestCases.SS.Util.Utils.AssertError(_fe, _testCell, formula, expectedError);
        }

        private ICell SetFormula(int rowIndex, int columnIndex, string formula)
        {
            IRow row = _sheet.GetRow(rowIndex) ?? _sheet.CreateRow(rowIndex);
            ICell cell = row.CreateCell(columnIndex);
            cell.SetCellFormula(formula);
            _fe.NotifyUpdateCell(cell);
            return cell;
        }

        private CellValue EvaluateFormulaAt(int rowIndex, int columnIndex, string formula)
        {
            ICell cell = SetFormula(rowIndex, columnIndex, formula);
            return _fe.Evaluate(cell);
        }

        // --- SEQUENCE ---

        [Test]
        public void TestSequence()
        {
            // the anchor formula itself evaluates to the first element of the array
            AssertNumber("SEQUENCE(3)", 1);
            // the whole array is available in array context
            AssertNumber("SUM(SEQUENCE(3))", 6);
            AssertNumber("SUM(SEQUENCE(2,3))", 21);
            AssertNumber("SUM(SEQUENCE(2,2,5,10))", 5 + 15 + 25 + 35);
            // width defaults to one
            AssertNumber("SUM(SEQUENCE(3,1))", 6);
            // single element result
            AssertNumber("SUM(SEQUENCE(1))", 1);
        }

        [Test]
        public void TestSequenceArgErrors()
        {
            AssertError("SEQUENCE(0)", FormulaError.VALUE);
            AssertError("SEQUENCE(3,0)", FormulaError.VALUE);
            AssertError("SEQUENCE(\"a\")", FormulaError.VALUE);
            // SEQUENCE takes at least one argument
            AssertError("SEQUENCE()", FormulaError.VALUE);
            // ... and at most four
            AssertError("SEQUENCE(1,1,1,1,1)", FormulaError.VALUE);
        }

        // --- RANDARRAY ---

        [Test]
        public void TestRandArray()
        {
            // bounds are validated
            AssertError("RANDARRAY(0)", FormulaError.VALUE);
            AssertError("RANDARRAY(1,0)", FormulaError.VALUE);
            AssertError("RANDARRAY(2,1,5,4)", FormulaError.VALUE);
            AssertError("RANDARRAY(1,1,5,4)", FormulaError.VALUE);
            // whole numbers are in [min, max)
            AssertNumber("RANDARRAY(2,2,1,10,TRUE)", 1, 9);
            // the single element of a 1x1 array is the value itself, in [0, 1)
            AssertNumber("RANDARRAY()", 0, 1);
        }

        [Test]
        public void TestRandArrayRandomness()
        {
            // repeated evaluations of RANDARRAY must produce different values
            _testCell.SetCellFormula("RANDARRAY()");
            _fe.NotifyUpdateCell(_testCell);
            double first = _fe.Evaluate(_testCell).NumberValue;
            bool different = false;
            for (int i = 0; i < 100 && !different; i++)
            {
                _testCell.SetCellFormula("RANDARRAY()");
                _fe.NotifyUpdateCell(_testCell);
                if (Math.Abs(_fe.Evaluate(_testCell).NumberValue - first) > 1e-9)
                {
                    different = true;
                }
            }
            ClassicAssert.IsTrue(different, "RANDARRAY should produce random values");
        }

        // --- SINGLE ('@') ---

        [Test]
        public void TestSingleOperatorImplicitIntersection()
        {
            TestCases.SS.Util.Utils.AddRow(_sheet, 0, 1, 2, 3);   // A1:C1
            TestCases.SS.Util.Utils.AddRow(_sheet, 1, 4, 5, 6);   // A2:C2
            TestCases.SS.Util.Utils.AddRow(_sheet, 2, 7, 8, 9);   // A3:C3

            // single row area: implicit intersection picks the column of the calling cell
            CellValue b2 = EvaluateFormulaAt(1, 1, "@A1:C1");
            ClassicAssert.AreEqual(CellType.Numeric, b2.CellType);
            ClassicAssert.AreEqual(2.0, b2.NumberValue);

            // single column area: implicit intersection picks the row of the calling cell
            CellValue d1 = EvaluateFormulaAt(0, 3, "@A1:A3");
            ClassicAssert.AreEqual(1.0, d1.NumberValue);
            CellValue c3 = EvaluateFormulaAt(2, 2, "@A1:A3");
            ClassicAssert.AreEqual(7.0, c3.NumberValue);

            // single cell operand is returned as is
            CellValue f1 = EvaluateFormulaAt(0, 5, "@A1");
            ClassicAssert.AreEqual(1.0, f1.NumberValue);

            // a calling position that does not intersect the area is #VALUE!
            CellValue e1 = EvaluateFormulaAt(0, 4, "_xlfn.SINGLE(A1:C1)");
            ClassicAssert.AreEqual(CellType.Error, e1.CellType);
            ClassicAssert.AreEqual(FormulaError.VALUE.Code, e1.ErrorValue);
            AssertError("@A1:C1", FormulaError.VALUE);

            // value operand is passed through
            AssertNumber("@5+1", 6);
        }

        // --- ANCHORARRAY ('#') ---

        [Test]
        public void TestSpilledRangeOperator()
        {
            TestCases.SS.Util.Utils.AddRow(_sheet, 0, 1, 2, 3);
            // anchor formula spilling a 3-cell array
            ICell anchor = SetFormula(1, 0, "SEQUENCE(3)");    // A2

            // the anchor itself evaluates to the first element
            ClassicAssert.AreEqual(1.0, _fe.Evaluate(anchor).NumberValue);

            // aggregate over the whole spill range
            AssertNumber("SUM(A2#)", 6);
            // the future-function form is equivalent to the operator
            AssertNumber("SUM(_xlfn.ANCHORARRAY(A2))", 6);
            // another dynamic array function as anchor
            SetFormula(1, 4, "SEQUENCE(2,3)");                 // E2
            AssertNumber("SUM(E2#)", 21);
            // spill ranges broadcast like arrays
            AssertNumber("SUMPRODUCT(A2#,A2#)", 14);
        }

        [Test]
        public void TestSpilledRangeOperatorNotAnchored()
        {
            TestCases.SS.Util.Utils.AddRow(_sheet, 0, 1, 2, 3);
            // '#' on a plain value cell is #REF!
            AssertError("SUM(A1#)", FormulaError.REF);
            // '#' on an empty cell
            AssertError("SUM(B10#)", FormulaError.REF);
            // a bare spill reference on a cell without an anchor is #REF! too
            AssertError("A2#", FormulaError.REF);
        }

        [Test]
        public void TestSpilledRangeOperatorCircular()
        {
            // an anchor formula that references its own spill range is a circular reference
            ICell anchor = SetFormula(1, 0, "SUM(A2#)");        // A2
            CellValue value = _fe.Evaluate(anchor);
            ClassicAssert.AreEqual(CellType.Error, value.CellType);
            ClassicAssert.AreEqual(ErrorEval.CIRCULAR_REF_ERROR.ErrorCode, value.ErrorValue);
        }

        [Test]
        public void TestSpilledRangeOperatorArrayFormulaGroup()
        {
            // legacy CSE array formula as the anchor of '#'
            CellRangeAddress range = new CellRangeAddress(1, 3, 0, 0);     // A2:A4
            _sheet.SetArrayFormula("SEQUENCE(3)", range);
            // the anchor evaluates to the first element of the group
            ICell anchor = _sheet.GetRow(1).GetCell(0);
            ClassicAssert.AreEqual(1.0, _fe.Evaluate(anchor).NumberValue);
            // and the spill reference aggregates the whole group
            AssertNumber("SUM(A2#)", 6);
            // a single cell outside of the group is not an anchor
            AssertError("SUM(B2#)", FormulaError.REF);
        }

        [Test]
        public void TestSpilledRangeOperatorElementWise()
        {
            // a 2x3 spill anchored at A2, referenced element-wise inside a legacy
            // array formula group of the same shape
            SetFormula(1, 0, "SEQUENCE(2,3)");                  // A2
            CellRangeAddress range = new CellRangeAddress(5, 6, 2, 4);     // C6:E7
            _sheet.SetArrayFormula("A2#", range);
            ClassicAssert.AreEqual(1.0, _fe.Evaluate(_sheet.GetRow(5).GetCell(2)).NumberValue, 0.0);
            ClassicAssert.AreEqual(2.0, _fe.Evaluate(_sheet.GetRow(5).GetCell(3)).NumberValue, 0.0);
            ClassicAssert.AreEqual(3.0, _fe.Evaluate(_sheet.GetRow(5).GetCell(4)).NumberValue, 0.0);
            ClassicAssert.AreEqual(4.0, _fe.Evaluate(_sheet.GetRow(6).GetCell(2)).NumberValue, 0.0);
            ClassicAssert.AreEqual(5.0, _fe.Evaluate(_sheet.GetRow(6).GetCell(3)).NumberValue, 0.0);
            ClassicAssert.AreEqual(6.0, _fe.Evaluate(_sheet.GetRow(6).GetCell(4)).NumberValue, 0.0);
        }

        // --- dynamic arrays combined ---

        [Test]
        public void TestDynamicArrayCombinations()
        {
            SetFormula(1, 0, "SEQUENCE(3)");                   // A2
            // spill ranges broadcast through array-aware functions
            AssertNumber("SUMPRODUCT(A2#,A2#)", 14);
            // a formula that consumes a spilled range and an array at once
            AssertNumber("SUM(SEQUENCE(3,1,1,1))", 6);
            // a cell consuming another cell's array result
            ICell e2 = SetFormula(1, 4, "SUM(SEQUENCE(3))");    // E2
            ClassicAssert.AreEqual(6.0, _fe.Evaluate(e2).NumberValue);
            AssertNumber("E2+1", 7);
        }
    }
}
