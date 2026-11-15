/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 *  ====================================================================
 */

using System.Collections.Generic;
using System.Xml;

using NPOI.OpenXmlFormats.Wordprocessing;
using NUnit.Framework;

namespace ooxml.Testcases
{
    /// <summary>
    /// Tests for the WordprocessingML form field (ffData) types, in particular the
    /// drop-down list form field used for legacy FORMDROPDOWN fields.
    /// </summary>
    [TestFixture]
    public class FormFieldTest
    {
        private static CT_FFData ParseFfData(string inner)
        {
            string xml = "<w:ffData xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                         + inner + "</w:ffData>";
            var doc = new XmlDocument();
            doc.LoadXml(xml);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            return CT_FFData.Parse(doc.DocumentElement, ns);
        }

        [Test]
        public void DropDownListParsesSelectedEntry()
        {
            CT_FFData ffData = ParseFfData(
                "<w:name w:val=\"Color\"/><w:enabled/>" +
                "<w:ddList><w:result w:val=\"1\"/><w:default w:val=\"0\"/>" +
                "<w:listEntry w:val=\"Red\"/><w:listEntry w:val=\"Green\"/><w:listEntry w:val=\"Blue\"/>" +
                "</w:ddList>");

            List<CT_FFDDList> dropDowns = ffData.GetDropDownList();
            Assert.AreEqual(1, dropDowns.Count);
            Assert.AreEqual("Green", dropDowns[0].GetSelectedValue());
        }

        [Test]
        public void DropDownListFallsBackToDefaultWhenResultMissing()
        {
            CT_FFData ffData = ParseFfData(
                "<w:ddList><w:default w:val=\"2\"/>" +
                "<w:listEntry w:val=\"Red\"/><w:listEntry w:val=\"Green\"/><w:listEntry w:val=\"Blue\"/>" +
                "</w:ddList>");

            Assert.AreEqual("Blue", ffData.GetDropDownList()[0].GetSelectedValue());
        }

        [Test]
        public void DropDownListFallsBackToDefaultWhenResultOutOfRange()
        {
            CT_FFData ffData = ParseFfData(
                "<w:ddList><w:result w:val=\"9\"/><w:default w:val=\"0\"/>" +
                "<w:listEntry w:val=\"Red\"/><w:listEntry w:val=\"Green\"/><w:listEntry w:val=\"Blue\"/>" +
                "</w:ddList>");

            Assert.AreEqual("Red", ffData.GetDropDownList()[0].GetSelectedValue());
        }

        [Test]
        public void DropDownListReturnsNullWhenNothingIsValid()
        {
            CT_FFData ffData = ParseFfData(
                "<w:ddList><w:result w:val=\"9\"/><w:default w:val=\"9\"/>" +
                "<w:listEntry w:val=\"Red\"/></w:ddList>");

            Assert.IsNull(ffData.GetDropDownList()[0].GetSelectedValue());
        }

        [Test]
        public void DropDownListRoundTripsThroughParser()
        {
            const string xml =
                "<w:ddList xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                "<w:result w:val=\"1\"/><w:default w:val=\"0\"/>" +
                "<w:listEntry w:val=\"Red\"/><w:listEntry w:val=\"Green\"/>" +
                "</w:ddList>";
            var doc = new XmlDocument();
            doc.LoadXml(xml);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            CT_FFDDList parsed = CT_FFDDList.Parse(doc.DocumentElement, ns);

            Assert.AreEqual("1", parsed.result.val);
            Assert.AreEqual("0", parsed.@default.val);
            Assert.AreEqual(new[] { "Red", "Green" }, parsed.listEntry.ConvertAll(e => e.val));
            Assert.AreEqual("Green", parsed.GetSelectedValue());
        }
    }
}
