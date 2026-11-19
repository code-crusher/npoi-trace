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

namespace TestCases.HSSF.Record
{
    using System;
    using System.IO;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using NPOI.Util;
    using TestCases.HSSF;
    using NUnit.Framework;using NUnit.Framework.Legacy;

    [TestFixture]
    public class TestXFExtRecord
    {
        [Test]
        public void TestParseAndSerialize()
        {
            // one foreground color extension: theme 5 (accent2), tint +26213
            XFExtRecord rec = new XFExtRecord();
            rec.Ixfe = 65;
            rec.SetColorExt(XFExtRecord.EXTPROP_INTERIOR_FOREGROUND,
                new FullColorExt(XFExtRecord.XCLRTHEMED, 26213, 5));

            byte[] data = rec.Serialize();
            XFExtRecord parsed = RoundTrip(rec);

            ClassicAssert.AreEqual(65, parsed.Ixfe);
            ClassicAssert.AreEqual(1, parsed.ExtProps.Count);

            FullColorExt fg = parsed.GetColorExt(XFExtRecord.EXTPROP_INTERIOR_FOREGROUND);
            ClassicAssert.IsNotNull(fg);
            ClassicAssert.AreEqual(XFExtRecord.XCLRTHEMED, fg.XclrType);
            ClassicAssert.AreEqual(26213, fg.TintShade);
            ClassicAssert.AreEqual(5, fg.XclrValue);
            ClassicAssert.AreEqual(data, parsed.Serialize());
        }

        [Test]
        public void TestUnknownExtPropsArePreserved()
        {
            XFExtRecord rec = new XFExtRecord();
            rec.Ixfe = 7;
            // an extension type we do not interpret (text color, 0x000D)
            byte[] raw = { 0x02, 0x00, 0xFF, 0x00, 0x00, 0x00 };
            rec.SetExtPropData(0x000D, raw);
            rec.SetColorExt(XFExtRecord.EXTPROP_INTERIOR_BACKGROUND,
                new FullColorExt(XFExtRecord.XCLRRGB, 0, 0x0000FF00));

            XFExtRecord parsed = RoundTrip(rec);
            ClassicAssert.AreEqual(2, parsed.ExtProps.Count);
            ClassicAssert.AreEqual(raw, parsed.GetExtPropData(0x000D));
            FullColorExt bg = parsed.GetColorExt(XFExtRecord.EXTPROP_INTERIOR_BACKGROUND);
            ClassicAssert.IsNotNull(bg);
            ClassicAssert.AreEqual(XFExtRecord.XCLRRGB, bg.XclrType);
            ClassicAssert.AreEqual(0x0000FF00, bg.XclrValue);
        }

        [Test]
        public void TestNegativeTint()
        {
            FullColorExt color = new FullColorExt(XFExtRecord.XCLRTHEMED, -16383, 0);
            XFExtRecord rec = new XFExtRecord();
            rec.Ixfe = 0;
            rec.SetColorExt(XFExtRecord.EXTPROP_INTERIOR_FOREGROUND, color);

            FullColorExt parsed = RoundTrip(rec).GetColorExt(XFExtRecord.EXTPROP_INTERIOR_FOREGROUND);
            ClassicAssert.IsNotNull(parsed);
            ClassicAssert.AreEqual(-16383, parsed.TintShade);
            ClassicAssert.AreEqual(0, parsed.XclrValue);
        }

        [Test]
        public void TestSetExtPropDataReplaces()
        {
            XFExtRecord rec = new XFExtRecord();
            rec.SetColorExt(XFExtRecord.EXTPROP_INTERIOR_FOREGROUND,
                new FullColorExt(XFExtRecord.XCLRAUTO, 0, 0));
            rec.SetColorExt(XFExtRecord.EXTPROP_INTERIOR_FOREGROUND,
                new FullColorExt(XFExtRecord.XCLRRGB, 0, 0x00FF0000));

            ClassicAssert.AreEqual(1, rec.ExtProps.Count);
            FullColorExt fg = rec.GetColorExt(XFExtRecord.EXTPROP_INTERIOR_FOREGROUND);
            ClassicAssert.AreEqual(XFExtRecord.XCLRRGB, fg.XclrType);
            ClassicAssert.AreEqual(0x00FF0000, fg.XclrValue);
        }

        [Test]
        public void TestRealFileRecords()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ColorTable.xls");
            XFExtRecord ext = wb.Workbook.GetXFExtRecord(75);
            ClassicAssert.IsNotNull(ext);
            FullColorExt fg = ext.GetColorExt(XFExtRecord.EXTPROP_INTERIOR_FOREGROUND);
            ClassicAssert.IsNotNull(fg);
            // cell B2 of the sample file: dk2 (theme 3) lightened by 80%
            ClassicAssert.AreEqual(XFExtRecord.XCLRTHEMED, fg.XclrType);
            ClassicAssert.AreEqual(3, fg.XclrValue);
            ClassicAssert.AreEqual(26213, fg.TintShade);
        }

        private static XFExtRecord RoundTrip(XFExtRecord rec)
        {
            byte[] data = rec.Serialize();
            MemoryStream stream = new MemoryStream(data);
            RecordInputStream in1 = new RecordInputStream(stream);
            in1.NextRecord();
            ClassicAssert.AreEqual(XFExtRecord.sid, in1.Sid);
            return new XFExtRecord(in1);
        }
    }
}
