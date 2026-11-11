using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.XSLF.UserModel;
using NUnit.Framework;

namespace TestCases.XSLF.UserModel
{
    [TestFixture]
    public class TestXSLFCharts
    {
        [Test]
        public void ChartPartsAreListedAndBytesArePreserved()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var charts = presentation.GetCharts();
            Assert.That(charts.Count, Is.EqualTo(1));
            Assert.That(charts[0].GetPackagePart().PartName.Name, Is.EqualTo("/ppt/charts/chart1.xml"));
            Assert.That(charts[0].GetPackagePart().ContentType,
                Is.EqualTo(XSLFRelation.CHART.ContentType));

            var chartBytes = ReadBytes(charts[0].GetPackagePart());
            var workbookPart = presentation.Package.GetPart(
                PackagingUriHelper.CreatePartName("/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"));
            var workbookBytes = ReadBytes(workbookPart);

            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetCharts().Count, Is.EqualTo(1));
            Assert.That(ReadBytes(reopened.GetCharts()[0].GetPackagePart()), Is.EqualTo(chartBytes));
            var savedWorkbook = reopened.Package.GetPart(
                PackagingUriHelper.CreatePartName("/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"));
            Assert.That(ReadBytes(savedWorkbook), Is.EqualTo(workbookBytes));
            using var again = Open(Save(reopened));
            Assert.That(ReadBytes(again.GetCharts()[0].GetPackagePart()), Is.EqualTo(chartBytes));
        }

        [Test]
        public void SnapshotsAreDetachedAndDoNotChangeThePackage()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var chart = presentation.GetCharts()[0];
            var originalXml = ReadXml(chart.GetPackagePart()).OuterXml;
            var originalBytes = ReadBytes(chart.GetPackagePart());

            var xmlSnapshot = chart.GetXmlSnapshot();
            xmlSnapshot.DocumentElement.SetAttribute("detached", "yes");
            var modelSnapshot = chart.GetChartSnapshot();
            Assert.That(modelSnapshot.chart, Is.Not.Null);
            modelSnapshot.chart.autoTitleDeleted = null;
            Assert.That(modelSnapshot.lang, Is.Not.Null);

            using var reopened = Open(Save(presentation));
            Assert.That(ReadXml(reopened.GetCharts()[0].GetPackagePart()).OuterXml, Is.EqualTo(originalXml));
            using var again = Open(Save(reopened));
            Assert.That(ReadBytes(again.GetCharts()[0].GetPackagePart()), Is.EqualTo(originalBytes));
        }

        [Test]
        public void ExistingChartFramesAreRecognized()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var slide = presentation.GetSlides()[0];
            var frames = slide.GetChartFrames();
            Assert.That(frames.Count, Is.EqualTo(1));
            Assert.That(frames[0].Name, Is.EqualTo("Chart 1"));
            Assert.That(frames[0].Anchor, Is.EqualTo(new XSLFChartAnchor(1524000, 1397000, 6096000, 4064000)));
        }

        [Test]
        public void ChartCanBeReusedOnAnotherSlideWithIndependentAnchorAndName()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var chart = presentation.GetCharts()[0];
            var originalChartBytes = ReadBytes(chart.GetPackagePart());
            var second = presentation.CreateSlide();
            var frame = second.AddChartFrame(chart);
            frame.Name = "Sales Chart";
            frame.Anchor = new XSLFChartAnchor(500000, 600000, 4000000, 3000000);

            using var reopened = Open(Save(presentation));
            var chartBytesAfter = ReadBytes(reopened.GetCharts()[0].GetPackagePart());
            Assert.That(chartBytesAfter, Is.EqualTo(originalChartBytes));
            var savedSlide = reopened.GetSlides()[1];
            var savedFrames = savedSlide.GetChartFrames();
            Assert.That(savedFrames.Count, Is.EqualTo(1));
            Assert.That(savedFrames[0].Name, Is.EqualTo("Sales Chart"));
            Assert.That(savedFrames[0].Anchor, Is.EqualTo(new XSLFChartAnchor(500000, 600000, 4000000, 3000000)));
            var firstSlideFrames = reopened.GetSlides()[0].GetChartFrames();
            Assert.That(firstSlideFrames.Count, Is.EqualTo(1));
            Assert.That(firstSlideFrames[0].Name, Is.EqualTo("Chart 1"));

            using var again = Open(Save(reopened));
            var persisted = again.GetSlides()[1].GetChartFrames();
            Assert.That(persisted[0].Name, Is.EqualTo("Sales Chart"));
            Assert.That(persisted[0].Anchor, Is.EqualTo(new XSLFChartAnchor(500000, 600000, 4000000, 3000000)));
        }

        [Test]
        public void SameChartCanBeReusedOnSeveralSheetsAndWithinOneSheet()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var chart = presentation.GetCharts()[0];
            var slide = presentation.GetSlides()[0];
            var layout = slide.GetSlideLayout();
            var master = presentation.GetSlideMasters()[0];

            var slideFrame = slide.AddChartFrame(chart);
            var layoutFrame = layout.AddChartFrame(chart);
            var masterFrame = master.AddChartFrame(chart);
            var secondSlideFrame = presentation.CreateSlide().AddChartFrame(chart);

            Assert.That(slideFrame, Is.Not.SameAs(layoutFrame));
            var slideFrames = slide.GetChartFrames();
            Assert.That(slideFrames.Count, Is.EqualTo(2));

            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetCharts().Count, Is.EqualTo(1));
            Assert.That(reopened.GetSlides()[0].GetChartFrames().Count, Is.EqualTo(2));
            Assert.That(reopened.GetSlides()[0].GetSlideLayout().GetChartFrames().Count, Is.EqualTo(1));
            Assert.That(reopened.GetSlideMasters()[0].GetChartFrames().Count, Is.EqualTo(1));
            Assert.That(reopened.GetSlides()[1].GetChartFrames().Count, Is.EqualTo(1));
            using var again = Open(Save(reopened));
            Assert.That(again.GetCharts().Count, Is.EqualTo(1));
            Assert.That(again.GetSlides()[0].GetChartFrames().Count, Is.EqualTo(2));
        }

        [Test]
        public void ReusedFrameShapeIdsAreSheetWideUnique()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var slide = presentation.GetSlides()[0];
            var chart = presentation.GetCharts()[0];
            slide.AddChartFrame(chart);
            slide.AddChartFrame(chart);

            using var reopened = Open(Save(presentation));
            var savedSlide = reopened.GetSlides()[0];
            var ids = ReadFrameIds(savedSlide);
            CollectionAssert.AllItemsAreUnique(ids);
            Assert.That(savedSlide.GetChartFrames().Count, Is.EqualTo(3));
            using var again = Open(Save(reopened));
            CollectionAssert.AllItemsAreUnique(ReadFrameIds(again.GetSlides()[0]));
        }

        [Test]
        public void RelationshipsAreReusedByTypeAndTarget()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var slide = presentation.GetSlides()[0];
            var chart = presentation.GetCharts()[0];
            var before = slide.GetPackagePart().GetRelationshipsByType(XSLFRelation.CHART.Relation).Size;
            Assert.That(before, Is.EqualTo(1));
            slide.AddChartFrame(chart);
            slide.AddChartFrame(chart);
            var after = slide.GetPackagePart().GetRelationshipsByType(XSLFRelation.CHART.Relation).Size;
            Assert.That(after, Is.EqualTo(1));
        }

        [Test]
        public void FrameNameAndAnchorEditingDoesNotTouchChartContents()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var chart = presentation.GetCharts()[0];
            var originalChartBytes = ReadBytes(chart.GetPackagePart());
            var frame = presentation.GetSlides()[0].GetChartFrames()[0];
            frame.Name = "Renamed Chart";
            frame.Anchor = new XSLFChartAnchor(1000, 2000, 3000, 4000);

            using var reopened = Open(Save(presentation));
            Assert.That(ReadBytes(reopened.GetCharts()[0].GetPackagePart()), Is.EqualTo(originalChartBytes));
            var savedFrame = reopened.GetSlides()[0].GetChartFrames()[0];
            Assert.That(savedFrame.Name, Is.EqualTo("Renamed Chart"));
            Assert.That(savedFrame.Anchor, Is.EqualTo(new XSLFChartAnchor(1000, 2000, 3000, 4000)));
            using var again = Open(Save(reopened));
            Assert.That(ReadBytes(again.GetCharts()[0].GetPackagePart()), Is.EqualTo(originalChartBytes));
            Assert.That(again.GetSlides()[0].GetChartFrames()[0].Name, Is.EqualTo("Renamed Chart"));
            Assert.That(again.GetSlides()[0].GetChartFrames()[0].Anchor,
                Is.EqualTo(new XSLFChartAnchor(1000, 2000, 3000, 4000)));
        }

        [Test]
        public void ForeignChartsAreRejected()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            using var otherStream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var other = new XMLSlideShow(otherStream);
            var slide = presentation.GetSlides()[0];
            var foreignChart = other.GetCharts()[0];
            Assert.Throws<ArgumentException>(() => slide.AddChartFrame(foreignChart));
        }

        [Test]
        public void RemovedChartPartsAreNotListed()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var chartPart = presentation.GetCharts()[0].GetPackagePart();
            presentation.Package.RemovePart(chartPart);
            Assert.That(presentation.GetCharts(), Is.Empty);
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides()[0].GetChartFrames().Count, Is.EqualTo(0));
        }

        [Test]
        public void NonChartGraphicFramesAreNotListed()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("shapes.pptx");
            using var presentation = new XMLSlideShow(stream);
            Assert.That(presentation.GetCharts(), Is.Empty);
            foreach (var slide in presentation.GetSlides())
            {
                Assert.That(slide.GetChartFrames(), Is.Empty);
            }
        }

        [Test]
        public void MalformedChartReferenceIsNotListedAsChartFrame()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var slide = presentation.GetSlides()[0];
            Assert.That(slide.GetChartFrames().Count, Is.EqualTo(1));
            var commonSlideData = slide.GetXmlObject().cSld;
            var manager = new XmlNamespaceManager(commonSlideData.OwnerDocument.NameTable);
            manager.AddNamespace("p", CT_Presentation.Namespace);
            manager.AddNamespace("a", XSLFXmlHelper.DrawingNamespace);
            manager.AddNamespace("c", XSLFChart.ChartNamespace);
            manager.AddNamespace("r", CT_Presentation.RelationshipsNamespace);
            var reference = (XmlElement)commonSlideData.SelectSingleNode(
                "//p:graphicFrame/a:graphic/a:graphicData/c:chart", manager);
            reference.RemoveAttribute("id", CT_Presentation.RelationshipsNamespace);

            Assert.That(slide.GetChartFrames(), Is.Empty);
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides()[0].GetChartFrames(), Is.Empty);
        }

        [Test]
        public void ReadOnlyAndClosedPresentationsCannotAddChartFrames()
        {
            byte[] bytes;
            using (var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx"))
            using (var presentation = new XMLSlideShow(stream))
            {
                bytes = Save(presentation);
            }
            using var readOnlyPackage = OPCPackage.Open(new MemoryStream(bytes), true);
            using var readOnly = new XMLSlideShow(readOnlyPackage);
            Assert.That(readOnly.GetCharts().Count, Is.EqualTo(1));
            Assert.Throws<InvalidOperationException>(() =>
                readOnly.GetSlides()[0].AddChartFrame(readOnly.GetCharts()[0]));

            var presentation2 = new XMLSlideShow(new MemoryStream(bytes));
            var closedChart = presentation2.GetCharts()[0];
            presentation2.Dispose();
            Assert.Throws<ObjectDisposedException>(() =>
                presentation2.GetSlides()[0].AddChartFrame(closedChart));
        }

        [Test]
        public void InvalidOperationsDoNotChangeThePackage()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var slide = presentation.GetSlides()[0];
            var partCount = presentation.Package.GetParts().Count;
            var relationshipCount = slide.GetPackagePart().GetRelationshipsByType(XSLFRelation.CHART.Relation).Size;
            var originalXml = ReadXml(slide.GetPackagePart()).OuterXml;

            Assert.Throws<ArgumentNullException>(() => slide.AddChartFrame(null));
            using var otherStream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var other = new XMLSlideShow(otherStream);
            Assert.Throws<ArgumentException>(() => slide.AddChartFrame(other.GetCharts()[0]));

            Assert.That(presentation.Package.GetParts().Count, Is.EqualTo(partCount));
            Assert.That(slide.GetPackagePart().GetRelationshipsByType(XSLFRelation.CHART.Relation).Size,
                Is.EqualTo(relationshipCount));
            Assert.That(ReadXml(slide.GetPackagePart()).OuterXml, Is.EqualTo(originalXml));
            using var reopened = Open(Save(presentation));
            AssertXmlEquivalent(originalXml, ReadXml(reopened.GetSlides()[0].GetPackagePart()).OuterXml);
        }

        [Test]
        public void InvalidFrameNameIsRejected()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var frame = presentation.GetSlides()[0].GetChartFrames()[0];
            Assert.Throws<ArgumentException>(() => frame.Name = null);
            Assert.Throws<ArgumentException>(() => frame.Name = string.Empty);
        }

        [Test]
        public void ChartFramesInsideNestedGroupsAreRecognized()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("pie-chart.pptx");
            using var presentation = new XMLSlideShow(stream);
            var slide = presentation.GetSlides()[0];
            var chart = presentation.GetCharts()[0];
            var tree = slide.GetXmlObject().cSld["spTree", CT_Presentation.Namespace];
            var group = XSLFXmlHelper.CreateElement(
                "<p:grpSp xmlns:p=\"" + CT_Presentation.Namespace + "\" xmlns:a=\"" +
                XSLFXmlHelper.DrawingNamespace + "\">" +
                "<p:nvGrpSpPr><p:cNvPr id=\"9\" name=\"Group 9\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
                "<p:grpSpPr/></p:grpSp>");
            group = (XmlElement)tree.AppendChild(tree.OwnerDocument.ImportNode(group, true));
            slide.AddChartFrame(chart);
            var frames = tree.ChildNodes.OfType<XmlElement>()
                .Where(element => element.LocalName == "graphicFrame").ToList();
            group.AppendChild(frames[frames.Count - 1]);

            using var reopened = Open(Save(presentation));
            var savedSlide = reopened.GetSlides()[0];
            Assert.That(savedSlide.GetChartFrames().Count, Is.EqualTo(2));
            Assert.That(ReadFrameIds(savedSlide), Is.Unique);
            using var again = Open(Save(reopened));
            Assert.That(again.GetSlides()[0].GetChartFrames().Count, Is.EqualTo(2));
        }

        private static List<uint> ReadFrameIds(XSLFSlide slide)
        {
            var document = ReadXml(slide.GetPackagePart());
            var tree = document.DocumentElement["cSld", CT_Presentation.Namespace]["spTree", CT_Presentation.Namespace];
            return tree.ChildNodes.OfType<XmlElement>()
                .Where(element => element.LocalName == "graphicFrame")
                .Select(element => uint.Parse(element["nvGraphicFramePr", CT_Presentation.Namespace]
                    ["cNvPr", CT_Presentation.Namespace].GetAttribute("id")))
                .ToList();
        }

        private static XMLSlideShow Open(byte[] bytes)
        {
            return new XMLSlideShow(new MemoryStream(bytes));
        }

        private static byte[] Save(XMLSlideShow presentation)
        {
            using var output = new MemoryStream();
            presentation.Write(output);
            return output.ToArray();
        }

        private static XmlDocument ReadXml(PackagePart part)
        {
            using var input = part.GetInputStream();
            var document = new XmlDocument { XmlResolver = null };
            document.Load(input);
            return document;
        }

        private static byte[] ReadBytes(PackagePart part)
        {
            using var input = part.GetInputStream();
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }

        private static void AssertXmlEquivalent(string expected, string actual)
        {
            var first = XElement.Parse(expected);
            var second = XElement.Parse(actual);
            first.DescendantsAndSelf().Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
            second.DescendantsAndSelf().Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
            Assert.That(XNode.DeepEquals(first, second), Is.True, "PresentationML content changed.");
        }
    }
}
