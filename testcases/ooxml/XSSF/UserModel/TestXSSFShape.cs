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

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.OpenXmlFormats.Vml;
using NPOI.OpenXmlFormats.Vml.Spreadsheet;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    internal class TestXSSFShape
    {
        [Test]
        public void Test58325_one()
        {
            Check58325(XSSFTestDataSamples.OpenSampleWorkbook("58325_lt.xlsx"), 1);
        }
        [Test]
        public void Test58325_three()
        {
            Check58325(XSSFTestDataSamples.OpenSampleWorkbook("58325_db.xlsx"), 3);
        }

        [Test]
        public void Test58325_three_RoundTripPreservesOleVml()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("58325_db.xlsx");
            XSSFWorkbook rt = XSSFTestDataSamples.WriteOutAndReadBack(wb);

            XSSFSheet sheet = rt.GetSheet("MetasNM001") as XSSFSheet;
            ClassicAssert.IsNotNull(sheet);

            // the OLE objects must still be registered on the sheet
            ClassicAssert.IsTrue(sheet.GetCTWorksheet().IsSetOleObjects());
            Assert.AreEqual(3, sheet.GetCTWorksheet().oleObjects.oleObject.Count);

            XSSFVMLDrawing vml = sheet.GetVMLDrawing(false);
            ClassicAssert.IsNotNull(vml, "VML drawing must survive a round-trip");

            // the picture/OLE shapetype must keep its id and o:spt so that
            // the shapes' type="#_x0000_t75" references stay intact
            ArrayList items = vml.GetItems();
            CT_Shapetype shapetype = null;
            var oleShapes = new List<CT_Shape>();
            foreach (object item in items)
            {
                if (item is CT_Shapetype st)
                    shapetype = st;
                else if (item is CT_Shape shape
                    && shape.sizeOfClientDataArray() > 0
                    && shape.GetClientDataArray(0).ObjectType == NPOI.OpenXmlFormats.Vml.Spreadsheet.ST_ObjectType.Pict)
                    oleShapes.Add(shape);
            }
            ClassicAssert.IsNotNull(shapetype);
            ClassicAssert.AreEqual("_x0000_t75", shapetype.id);
            ClassicAssert.AreEqual(75.0f, shapetype.spt);

            ClassicAssert.AreEqual(3, oleShapes.Count, "the 3 embedded OLE objects must keep their VML shapes");
            foreach (CT_Shape shape in oleShapes)
            {
                ClassicAssert.AreEqual("#_x0000_t75", shape.type);
                // the shape explicitly overrides the shapetype's filled="f"
                ClassicAssert.IsTrue(shape.filledSpecified, "filled must survive");
                ClassicAssert.AreNotEqual(NPOI.OpenXmlFormats.Vml.ST_TrueFalse.f, shape.filled);
                ClassicAssert.AreEqual("windowText [64]", shape.strokecolor);

                CT_ClientData cldata = shape.GetClientDataArray(0);
                // x:CF/x:AutoPict mark the shape as an OLE picture; losing them
                // degrades the embedded object to a plain image (issue #845)
                ClassicAssert.IsTrue(cldata.sizeWithCellsSpecified, "x:SizeWithCells must survive");
                ClassicAssert.AreEqual(NPOI.OpenXmlFormats.Vml.Spreadsheet.ST_CF.Pict, cldata.cf);
                ClassicAssert.IsTrue(cldata.autoPictSpecified, "x:AutoPict must survive");
                ClassicAssert.IsNotNull(cldata.anchor);
            }

            // the embedded binaries themselves must survive:
            // 3 OLE objects on the first sheet plus an embedded workbook on the second
            ClassicAssert.AreEqual(4, rt.GetAllEmbedds().Count);

            rt.Close();
            wb.Close();
        }

        private void Check58325(XSSFWorkbook wb, int expectedShapes)
        {
            XSSFSheet sheet = wb.GetSheet("MetasNM001") as XSSFSheet;
            ClassicAssert.IsNotNull(sheet);
            StringBuilder str = new StringBuilder();
            str.Append("sheet " + sheet.SheetName + " - ");
            XSSFDrawing drawing = sheet.GetDrawingPatriarch();
            //drawing = ((XSSFSheet)sheet).createDrawingPatriarch();
            List<XSSFShape> shapes = drawing.GetShapes();
            str.Append("drawing.Shapes.size() = " + shapes.Count);
            IEnumerator<XSSFShape> it = shapes.GetEnumerator();
            while (it.MoveNext())
            {
                XSSFShape shape = it.Current;
                str.Append(", " + shape.ToString());
                str.Append(", Col1:" + ((XSSFClientAnchor)shape.GetAnchor()).Col1);
                str.Append(", Col2:" + ((XSSFClientAnchor)shape.GetAnchor()).Col2);
                str.Append(", Row1:" + ((XSSFClientAnchor)shape.GetAnchor()).Row1);
                str.Append(", Row2:" + ((XSSFClientAnchor)shape.GetAnchor()).Row2);
            }

            ClassicAssert.AreEqual(expectedShapes, shapes.Count,
                "Having shapes: " + str);
        }
    }
}