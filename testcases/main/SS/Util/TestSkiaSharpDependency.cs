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

namespace TestCases.SS.Util
{
    using System;
    using System.IO;
    using System.Reflection;

    using SkiaSharp;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    using TestCases;

    /// <summary>
    /// Guards against issue #1877 / #1848: SkiaSharp must resolve and load in a
    /// consuming process. NPOI 2.8.0/2.8.1 packages shipped nuspec metadata that
    /// excluded SkiaSharp's runtime assets from consumers, so any NPOI code path
    /// that measures text or decodes images (AutoSizeColumn, GetImageDimension)
    /// crashed with FileNotFoundException for the SkiaSharp assembly.
    /// </summary>
    [TestFixture]
    public class TestSkiaSharpDependency
    {
        [Test]
        public void TestSkiaSharpAssemblyResolvesFromNpoiCoreLocation()
        {
            // Resolve SkiaSharp the way the runtime would when NPOI.Core needs it:
            // by name only, from the directory that hosts the NPOI assemblies.
            AssemblyName skiaName = new AssemblyName("SkiaSharp");

            Assembly skiaSharp = null;
            Assert.DoesNotThrow(() => skiaSharp = Assembly.Load(skiaName),
                "NPOI consumers must be able to load SkiaSharp by name (issue #1877)");

            ClassicAssert.IsNotNull(skiaSharp);
            ClassicAssert.AreEqual("SkiaSharp", skiaSharp.GetName().Name);
        }

        [Test]
        public void TestAutoSizeColumnUsesSkiaSharpTextMeasurement()
        {
            // AutoSizeColumn is the code path from the #1877/#1848 stack traces;
            // it measures text with SKFont and throws FileNotFound without SkiaSharp.
            using (IWorkbook wb = new HSSFWorkbook())
            {
                ISheet sheet = wb.CreateSheet("test");
                IRow row = sheet.CreateRow(0);
                row.CreateCell(0).SetCellValue("some text to measure");

                Assert.DoesNotThrow(() => sheet.AutoSizeColumn(0));
                ClassicAssert.Greater(sheet.GetColumnWidth(0), 0);

                ClassicAssert.Greater(SheetUtil.GetDefaultCharWidth(wb), 0);
            }
        }

        [Test]
        public void TestGetImageDimensionDecodesPngWithSkiaSharp()
        {
            using (FileStream png = POIDataSamples.GetSpreadSheetInstance().GetFile("45829.png"))
            {
                Assert.DoesNotThrow(() =>
                {
                    SKSizeI size = ImageUtils.GetImageDimension(png);
                    ClassicAssert.Greater(size.Width, 0);
                    ClassicAssert.Greater(size.Height, 0);
                });
            }
        }
    }
}
