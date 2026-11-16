/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace TestCases.XSSF.UserModel
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Xml.Linq;
    using System.Xml.XPath;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    /// <summary>
    /// Regression tests for issue 1201: a workbook template containing pivot
    /// tables together with slicers must survive a read/write round-trip so
    /// that Excel can open the result without reporting a repair.
    /// </summary>
    [TestFixture]
    public class TestXSSFPivotTableWithSlicersRoundtrip
    {
        private static XSSFWorkbook OpenSampleWorkbook()
        {
            return XSSFTestDataSamples.OpenSampleWorkbook("1201_TemplateWithSlicers.xlsx");
        }

        private static byte[] WriteOut(XSSFWorkbook wb)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                wb.Write(ms, true);
                return ms.ToArray();
            }
        }

        private static string ReadPart(XSSFWorkbook wb, string contentType)
        {
            foreach (PackagePart part in wb.Package.GetParts())
            {
                if (part.ContentType == contentType)
                {
                    using (StreamReader reader = new StreamReader(part.GetInputStream()))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        [Test]
        public void TestRoundtripPreservesSlicerDrawingNamespaces()
        {
            using (XSSFWorkbook wb = OpenSampleWorkbook())
            {
                byte[] output = WriteOut(wb);

                using (MemoryStream ms = new MemoryStream(output))
                using (XSSFWorkbook wb2 = new XSSFWorkbook(ms))
                {
                    ClassicAssert.AreEqual(wb.NumberOfSheets, wb2.NumberOfSheets);

                    string drawingXml = ReadPart(
                        wb2,
                        "application/vnd.openxmlformats-officedocument.drawing+xml");
                    ClassicAssert.IsNotNull(drawingXml, "the template's slicer drawing part must be preserved");

                    XDocument doc = XDocument.Parse(drawingXml);
                    int checkedChoices = 0;
                    foreach (XElement choice in doc.XPathSelectElements("//*[local-name()='Choice']"))
                    {
                        XAttribute requires = choice.Attribute("Requires");
                        if (requires == null)
                        {
                            continue;
                        }
                        checkedChoices++;
                        // the prefix referenced by mc:Requires must resolve to a namespace declaration
                        XElement alternateContent = choice.Parent;
                        ClassicAssert.IsNotNull(alternateContent);
                        bool declared = alternateContent.Attributes()
                            .Any(attr => attr.Name.LocalName == requires.Value);
                        ClassicAssert.IsTrue(declared,
                            "mc:Choice Requires='" + requires.Value + "' must have a matching xmlns declaration on mc:AlternateContent");
                    }
                    ClassicAssert.Greater(checkedChoices, 0,
                        "the drawing should contain mc:Choice elements referencing namespaces");
                }
            }
        }

        [Test]
        public void TestRoundtripKeepsPivotCacheDefinitionValid()
        {
            using (XSSFWorkbook wb = OpenSampleWorkbook())
            {
                byte[] output = WriteOut(wb);

                using (MemoryStream ms = new MemoryStream(output))
                using (XSSFWorkbook wb2 = new XSSFWorkbook(ms))
                {
                    string cacheXml = ReadPart(
                        wb2,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
                    ClassicAssert.IsNotNull(cacheXml, "pivotCacheDefinition part must be preserved");

                    XDocument doc = XDocument.Parse(cacheXml);
                    ClassicAssert.IsFalse(doc.Root.Attributes().Any(attr => attr.Name.LocalName == "tupleCache1"),
                        "pivotCacheDefinition must not contain the undeclared attribute tupleCache1");
                }
            }
        }
    }
}
