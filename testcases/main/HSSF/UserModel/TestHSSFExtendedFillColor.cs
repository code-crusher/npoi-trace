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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.Collections.Generic;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Util;
    using NPOI.SS.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;

    /**
     * Tests that fill colors stored as XFEXT color extensions (theme/RGB
     * colors instead of palette indexes) are reported correctly, i.e. cells
     * sharing a legacy palette index but displaying different colors get
     * different RGB values. See GitHub issue #820.
     */
    [TestFixture]
    public class TestHSSFExtendedFillColor
    {
        private static ICell Cell(HSSFWorkbook wb, int row, int col)
        {
            return wb.GetSheetAt(0).GetRow(row).GetCell(col);
        }

        /**
         * The sample file colors its cells with theme colors. Excel renders
         * these colors, not the stale palette indexes of the classic XF
         * records, and several cells share one palette index.
         */
        [Test]
        public void TestThemeColorsDifferDespiteSameLegacyIndex()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ColorTable.xls");

            // B2 (row 2, col 1) and B3 (row 3, col 1) both report legacy index 0x3F,
            // but Excel shows a light grey (accent5, +80%) and a dark grey
            // (lt1, -50%) respectively
            ICell b2 = Cell(wb, 2, 1);
            ICell b3 = Cell(wb, 3, 1);
            ClassicAssert.AreEqual(0x3F, b2.CellStyle.FillForegroundColorColor.Indexed);
            ClassicAssert.AreEqual(0x3F, b3.CellStyle.FillForegroundColorColor.Indexed);
            ClassicAssert.AreEqual(new byte[] { 89, 89, 89 }, b2.CellStyle.FillForegroundColorColor.RGB);
            ClassicAssert.AreEqual(new byte[] { 64, 64, 64 }, b3.CellStyle.FillForegroundColorColor.RGB);

            // A2 and C2 also share a legacy index (0x16) with different colors
            ICell a2 = Cell(wb, 2, 0);
            ICell c2 = Cell(wb, 2, 2);
            ClassicAssert.AreEqual(0x16, a2.CellStyle.FillForegroundColorColor.Indexed);
            ClassicAssert.AreEqual(0x16, c2.CellStyle.FillForegroundColorColor.Indexed);
            ClassicAssert.AreEqual(new byte[] { 217, 217, 217 }, a2.CellStyle.FillForegroundColorColor.RGB);
            ClassicAssert.AreEqual(new byte[] { 173, 173, 173 }, c2.CellStyle.FillForegroundColorColor.RGB);
        }

        /**
         * Row 0 of the sample file holds the plain theme colors. Their RGB
         * values must come from the workbook theme, not from the palette.
         */
        [Test]
        public void TestPlainThemeColors()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ColorTable.xls");

            ClassicAssert.AreEqual(new byte[] { 0, 0, 0 }, Cell(wb, 0, 1).CellStyle.FillForegroundColorColor.RGB);
            ClassicAssert.AreEqual(new byte[] { 68, 114, 196 }, Cell(wb, 0, 4).CellStyle.FillForegroundColorColor.RGB);
            ClassicAssert.AreEqual(new byte[] { 237, 125, 49 }, Cell(wb, 0, 5).CellStyle.FillForegroundColorColor.RGB);
            ClassicAssert.AreEqual(new byte[] { 255, 192, 0 }, Cell(wb, 0, 7).CellStyle.FillForegroundColorColor.RGB);
        }

        /**
         * Tinted variants of one theme color must be resolved per cell, e.g.
         * the yellow column of the sample file (accent4 = FFC000) darkened
         * by 40% renders as 255,217,102.
         */
        [Test]
        public void TestTintedThemeColor()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ColorTable.xls");

            // H1: accent4, +40%; H4: accent4, +20%
            ClassicAssert.AreEqual(new byte[] { 255, 242, 204 }, Cell(wb, 1, 7).CellStyle.FillForegroundColorColor.RGB);
            ClassicAssert.AreEqual(new byte[] { 255, 230, 153 }, Cell(wb, 2, 7).CellStyle.FillForegroundColorColor.RGB);
            ClassicAssert.AreEqual(new byte[] { 255, 217, 102 }, Cell(wb, 3, 7).CellStyle.FillForegroundColorColor.RGB);
        }

        /**
         * The reported color must survive a save/reload cycle, i.e. the
         * XFEXT records must be written back out.
         */
        [Test]
        public void TestRoundTripPreservesExtendedColors()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ColorTable.xls");
            byte[] rgbBefore = Cell(wb, 3, 7).CellStyle.FillForegroundColorColor.RGB;
            byte[] rgbBefore2 = Cell(wb, 2, 1).CellStyle.FillForegroundColorColor.RGB;

            HSSFWorkbook readBack = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            ClassicAssert.AreEqual(rgbBefore, Cell(readBack, 3, 7).CellStyle.FillForegroundColorColor.RGB);
            ClassicAssert.AreEqual(rgbBefore2, Cell(readBack, 2, 1).CellStyle.FillForegroundColorColor.RGB);
        }

        /**
         * Styles of workbooks without XFEXT records (the common case) keep
         * returning palette colors.
         */
        [Test]
        public void TestPaletteColorsUnaffected()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithColours.xls");
            ICell cell = wb.GetSheetAt(0).GetRow(1).GetCell(0);

            ClassicAssert.IsNotNull(cell.CellStyle.FillForegroundColorColor);
            ClassicAssert.AreEqual(typeof(HSSFColor), cell.CellStyle.FillForegroundColorColor.GetType().BaseType);
        }

        /**
         * A newly created workbook has no theme; its styles must keep working
         * through the palette.
         */
        [Test]
        public void TestNewWorkbook()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.CellStyle.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle.FillForegroundColor = HSSFColor.Red.Index;

            IColor color = cell.CellStyle.FillForegroundColorColor;
            ClassicAssert.IsNotNull(color);
            ClassicAssert.AreEqual(HSSFColor.Red.Index, color.Indexed);
            ClassicAssert.AreEqual(HSSFColor.Red.Triplet, color.RGB);
        }
    }
}
