using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.XSLF.UserModel;
using NPOI.XDDF.UserModel;
using NUnit.Framework;
using SkiaSharp;

namespace TestCases.XSLF.UserModel
{
    [TestFixture]
    public class TestXSLFShapeContainers
    {
        [Test]
        public void SheetShapeEnumerationUsesXmlOrder()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("shapes.pptx");
            using var presentation = new XMLSlideShow(input);
            var slide = presentation.GetSlides()[0];
            var shapes = slide.GetShapes();
            Assert.That(shapes.Select(shape => shape.GetType()), Is.EqualTo(new[]
            {
                typeof(XSLFSimpleShape), typeof(XSLFOpaqueShape), typeof(XSLFSimpleShape),
                typeof(XSLFOpaqueShape), typeof(XSLFOpaqueShape), typeof(XSLFOpaqueShape),
                typeof(XSLFOpaqueShape)
            }));
            Assert.That(shapes.Select(shape => shape.ShapeId),
                Is.EqualTo(new uint[] { 4, 6, 7, 2, 3, 8, 10 }));
            Assert.That(shapes.OfType<XSLFOpaqueShape>().Select(shape => shape.DrawingType),
                Is.EqualTo(new[] { "cxnSp", "pic", "graphicFrame", "cxnSp", "cxnSp" }));
            Assert.That(shapes.OfType<XSLFOpaqueShape>().Select(shape => shape.ShapeName),
                Is.EqualTo(new[]
                {
                    "Straight Connector 5", "Picture 1", "Table 2",
                    "Straight Arrow Connector 7", "Elbow Connector 9"
                }));
            Assert.That(((ICollection<XSLFShape>)shapes).IsReadOnly, Is.True);
            Assert.That(shapes[0].GetContainer(), Is.SameAs(slide));
        }

        [Test]
        public void GroupChildrenAreNestedAndEnumeratedInOrder()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("shapes.pptx");
            using var presentation = new XMLSlideShow(input);
            var group = (XSLFGroupShape)presentation.GetSlides()[2].GetShapes()[0];
            Assert.That(group.ShapeId, Is.EqualTo(5U));
            Assert.That(group.ShapeName, Is.EqualTo("Group 4"));
            Assert.That(group.GetShapes().Select(shape => shape.ShapeId), Is.EqualTo(new uint[] { 2, 3, 4 }));
            Assert.That(group.GetShapes().Select(shape => shape.GetContainer()), Has.All.SameAs(group));
        }

        [Test]
        public void CreatedSimpleShapesAndGroupsSurviveRepeatedSaveAndReopenCycles()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var rectangle = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            var group = slide.CreateGroup("Group");
            Assert.That(rectangle.ShapeId, Is.EqualTo(2U));
            Assert.That(group.ShapeId, Is.EqualTo(3U));
            Assert.That(slide.GetShapes().Select(shape => shape.ShapeId), Is.EqualTo(new uint[] { 2, 3 }));

            rectangle.Bounds = SKRect.Create(72, 36, 144, 72);
            group.Bounds = SKRect.Create(20, 40, 200, 100);
            group.InteriorBounds = SKRect.Create(0, 0, 100, 50);
            var ellipse = group.CreateSimpleShape(PresetGeometry.Ellipse, "Ellipse");
            var nested = group.CreateGroup("Nested");
            ellipse.Bounds = SKRect.Create(5, 5, 50, 25);
            nested.Bounds = SKRect.Create(1, 2, 10, 5);
            Assert.That(ellipse.ShapeId, Is.EqualTo(4U));
            Assert.That(nested.ShapeId, Is.EqualTo(5U));

            using var first = Open(Save(presentation));
            Save(first);
            using var second = Open(Save(first));
            var savedSlide = second.GetSlides()[0];
            var savedShapes = savedSlide.GetShapes();
            Assert.That(savedShapes.Select(shape => shape.ShapeId), Is.EqualTo(new uint[] { 2, 3 }));
            var savedRectangle = (XSLFSimpleShape)savedShapes[0];
            var savedGroup = (XSLFGroupShape)savedShapes[1];
            Assert.That(savedRectangle.ShapeName, Is.EqualTo("Rectangle"));
            Assert.That(savedRectangle.Geometry, Is.EqualTo(PresetGeometry.Rectangle));
            Assert.That(savedRectangle.Bounds, Is.EqualTo(SKRect.Create(72, 36, 144, 72)));
            Assert.That(savedGroup.Bounds, Is.EqualTo(SKRect.Create(20, 40, 200, 100)));
            Assert.That(savedGroup.InteriorBounds, Is.EqualTo(SKRect.Create(0, 0, 100, 50)));
            var savedChildren = savedGroup.GetShapes();
            Assert.That(savedChildren.Select(shape => shape.ShapeId), Is.EqualTo(new uint[] { 4, 5 }));
            var savedEllipse = (XSLFSimpleShape)savedChildren[0];
            var savedNested = (XSLFGroupShape)savedChildren[1];
            Assert.That(savedEllipse.ShapeName, Is.EqualTo("Ellipse"));
            Assert.That(savedEllipse.Geometry, Is.EqualTo(PresetGeometry.Ellipse));
            Assert.That(savedEllipse.Bounds, Is.EqualTo(SKRect.Create(5, 5, 50, 25)));
            Assert.That(savedNested.ShapeName, Is.EqualTo("Nested"));
            Assert.That(savedNested.Bounds, Is.EqualTo(SKRect.Create(1, 2, 10, 5)));

            // New identifiers stay stable and unique after appending more shapes.
            var appended = savedSlide.CreateSimpleShape(PresetGeometry.RightArrow, "Arrow");
            Assert.That(appended.ShapeId, Is.EqualTo(6U));
            var appendedGroup = savedGroup.CreateGroup("Appended");
            Assert.That(appendedGroup.ShapeId, Is.EqualTo(7U));
            using var third = Open(Save(second));
            Assert.That(third.GetSlides()[0].GetShapes().Select(shape => shape.ShapeId),
                Is.EqualTo(new uint[] { 2, 3, 6 }));
            var thirdGroup = (XSLFGroupShape)third.GetSlides()[0].GetShapes()[1];
            Assert.That(thirdGroup.GetShapes().Select(shape => shape.ShapeId),
                Is.EqualTo(new uint[] { 4, 5, 7 }));
        }

        [TestCase("sample.pptx")]
        [TestCase("shapes.pptx")]
        [TestCase("layouts.pptx")]
        [TestCase("themes.pptx")]
        [TestCase("pie-chart.pptx")]
        public void ExistingIdentifiersStayStableWhenShapesAreAppended(string name)
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream(name);
            using var presentation = new XMLSlideShow(stream);
            foreach (var sheet in Sheets(presentation))
            {
                var original = sheet.GetShapes().Select(shape => shape.ShapeId).ToArray();
                var added = sheet.CreateSimpleShape(PresetGeometry.Rectangle, "Added");
                Assert.That(added.ShapeId, Is.Not.EqualTo(0U));
                CollectionAssert.DoesNotContain(original, added.ShapeId);
                CollectionAssert.IsSubsetOf(original, sheet.GetShapes().Select(shape => shape.ShapeId));
            }
        }

        [Test]
        public void NewIdentifiersDoNotCollideWithIdentifiersInsideUnsupportedContent()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var tree = slide.GetXmlObject().cSld["spTree", CT_Presentation.Namespace];
            var document = tree.OwnerDocument;
            var picture = (XmlElement)document.ImportNode(XSLFXmlHelperProxy.CreateElement(
                "<p:pic xmlns:p=\"" + CT_Presentation.Namespace + "\">" +
                "<p:nvPicPr><p:cNvPr id=\"2\" name=\"Hidden Picture\"/><p:cNvPicPr/></p:nvPicPr>" +
                "<p:blipFill/><p:spPr/></p:pic>"), true);
            tree.AppendChild(picture);
            var extension = (XmlElement)document.ImportNode(XSLFXmlHelperProxy.CreateElement(
                "<p:extLst xmlns:p=\"" + CT_Presentation.Namespace + "\">" +
                "<p:ext uri=\"urn:npoi:test\"><p:cNvPr id=\"3\" name=\"\"/></p:ext></p:extLst>"), true);
            tree.AppendChild(extension);

            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Shape");
            var group = slide.CreateGroup("Group");
            var inner = group.CreateSimpleShape(PresetGeometry.Rectangle, "Inner");

            Assert.That(shape.ShapeId, Is.EqualTo(4U), "Ids 2 and 3 are used by unsupported content.");
            Assert.That(group.ShapeId, Is.EqualTo(5U));
            Assert.That(inner.ShapeId, Is.EqualTo(6U));
            using var reopened = Open(Save(presentation));
            Assert.That(reopened.GetSlides()[0].GetShapes().Select(s => s.ShapeId),
                Is.EqualTo(new uint[] { 2, 4, 5 }));
            var reopenedGroup = (XSLFGroupShape)reopened.GetSlides()[0].GetShapes()[2];
            Assert.That(reopenedGroup.GetShapes().Select(s => s.ShapeId), Is.EqualTo(new uint[] { 6 }));
        }

        [Test]
        public void PointAndEmuConversionAppliesNpoiUnits()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            shape.Bounds = SKRect.Create(72, 96.5f, 144, 288);
            var transform = FindTransform(shape);
            Assert.That(GetAttribute(transform, "off", "x"), Is.EqualTo("914400"));
            Assert.That(GetAttribute(transform, "off", "y"), Is.EqualTo("1225550"));
            Assert.That(GetAttribute(transform, "ext", "cx"), Is.EqualTo("1828800"));
            Assert.That(GetAttribute(transform, "ext", "cy"), Is.EqualTo("3657600"));
            Assert.That(shape.Bounds, Is.EqualTo(SKRect.Create(72, 96.5f, 144, 288)));
        }

        [Test]
        public void NegativePositionsAndBoundaryExtentsAreEditable()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            shape.Bounds = SKRect.Create(-72.5f, -1, 0, 0);
            Assert.That(shape.Bounds, Is.EqualTo(SKRect.Create(-72.5f, -1, 0, 0)));
            shape.Bounds = SKRect.Create(100000, 100000, 50000, 50000);
            Assert.That(shape.Bounds, Is.EqualTo(SKRect.Create(100000, 100000, 50000, 50000)));

            using var reopened = Open(Save(presentation));
            var saved = (XSLFSimpleShape)reopened.GetSlides()[0].GetShapes()[0];
            Assert.That(saved.Bounds, Is.EqualTo(SKRect.Create(100000, 100000, 50000, 50000)));
        }

        [TestCase(float.PositiveInfinity)]
        [TestCase(float.MinValue)]
        public void UnrepresentableCoordinatesAreRejectedWithoutMutation(float value)
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            shape.Bounds = SKRect.Create(10, 10, 10, 10);
            var before = shape.GetXmlObject().OuterXml;
            Assert.That(() => shape.Bounds = SKRect.Create(value, 0, 10, 10),
                Throws.TypeOf<POIXMLException>());
            Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
        }

        [TestCase(-1f)]
        [TestCase(-72f)]
        public void NegativeExtentsAreRejectedBeforeMutation(float width)
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            shape.Bounds = SKRect.Create(10, 10, 10, 10);
            var before = shape.GetXmlObject().OuterXml;
            Assert.That(() => shape.Bounds = SKRect.Create(10, 10, width, 10),
                Throws.TypeOf<ArgumentException>());
            Assert.That(shape.GetXmlObject().OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void GroupExteriorAndInteriorTransformsAreSeparate()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var group = slide.CreateGroup("Group");
            group.Bounds = SKRect.Create(100, 200, 300, 150);
            Assert.That(group.InteriorBounds, Is.EqualTo(SKRect.Create(0, 0, 0, 0)));
            group.InteriorBounds = SKRect.Create(0, 0, 600, 300);
            Assert.That(group.Bounds, Is.EqualTo(SKRect.Create(100, 200, 300, 150)));
            Assert.That(group.InteriorBounds, Is.EqualTo(SKRect.Create(0, 0, 600, 300)));

            var transform = FindTransform(group);
            Assert.That(GetAttribute(transform, "off", "x"), Is.EqualTo("1270000"));
            Assert.That(GetAttribute(transform, "ext", "cx"), Is.EqualTo("3810000"));
            Assert.That(GetAttribute(transform, "chOff", "x"), Is.EqualTo("0"));
            Assert.That(GetAttribute(transform, "chExt", "cx"), Is.EqualTo("7620000"));

            using var reopened = Open(Save(presentation));
            var saved = (XSLFGroupShape)reopened.GetSlides()[0].GetShapes()[0];
            Assert.That(saved.Bounds, Is.EqualTo(SKRect.Create(100, 200, 300, 150)));
            Assert.That(saved.InteriorBounds, Is.EqualTo(SKRect.Create(0, 0, 600, 300)));
        }

        [Test]
        public void GroupInteriorBoundsRejectNegativeExtentsWithoutMutation()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var group = slide.CreateGroup("Group");
            group.InteriorBounds = SKRect.Create(0, 0, 100, 100);
            var before = group.GetXmlObject().OuterXml;
            Assert.That(() => group.InteriorBounds = SKRect.Create(0, 0, -1, 100),
                Throws.TypeOf<ArgumentException>());
            Assert.That(group.GetXmlObject().OuterXml, Is.EqualTo(before));
        }

        [Test]
        public void RotationAndFlipTransformsSurviveBoundsEdits()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            var transform = FindTransform(shape);
            transform.SetAttribute("rot", "2700000");
            transform.SetAttribute("flipV", "1");
            var custom = transform.OwnerDocument.CreateElement("future", "ext", "urn:npoi:test");
            transform.AppendChild(custom);

            shape.Bounds = SKRect.Create(15, 25, 35, 45);

            using var reopened = Open(Save(presentation));
            var saved = (XSLFSimpleShape)reopened.GetSlides()[0].GetShapes()[0];
            var savedTransform = FindTransform(saved);
            Assert.That(savedTransform.GetAttribute("rot"), Is.EqualTo("2700000"));
            Assert.That(savedTransform.GetAttribute("flipV"), Is.EqualTo("1"));
            Assert.That(savedTransform["ext", "urn:npoi:test"], Is.Not.Null);
            Assert.That(saved.Bounds, Is.EqualTo(SKRect.Create(15, 25, 35, 45)));
        }

        [Test]
        public void UnsupportedContentAndRelationshipsSurviveEditingSupportedShapes()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("shapes.pptx");
            using var presentation = new XMLSlideShow(input);
            var slide = presentation.GetSlides()[0];
            var originalXml = ReadXml(slide.GetPackagePart()).OuterXml;
            var originalRelationships = RelationshipKeys(slide.GetPackagePart());

            var textBox = (XSLFSimpleShape)slide.GetShapes().First(shape => shape.ShapeId == 4);
            textBox.ShapeName = "Edited TextBox";
            textBox.Bounds = SKRect.Create(100, 100, 200, 50);

            using var reopened = Open(Save(presentation));
            var savedSlide = reopened.GetSlides()[0];
            var expected = XElement.Parse(originalXml);
            var edited = expected.Descendants(XName.Get("sp", CT_Presentation.Namespace)).First(sp =>
                sp.Descendants(XName.Get("cNvPr", CT_Presentation.Namespace))
                    .Any(pr => (string)pr.Attribute("id") == "4"));
            edited.Descendants(XName.Get("cNvPr", CT_Presentation.Namespace))
                .First().SetAttributeValue("name", "Edited TextBox");
            var editedTransform = edited.Descendants(XName.Get("xfrm", XSLFNamespace)).First();
            editedTransform.Element(XName.Get("off", XSLFNamespace)).SetAttributeValue("x", "1270000");
            editedTransform.Element(XName.Get("off", XSLFNamespace)).SetAttributeValue("y", "1270000");
            editedTransform.Element(XName.Get("ext", XSLFNamespace)).SetAttributeValue("cx", "2540000");
            editedTransform.Element(XName.Get("ext", XSLFNamespace)).SetAttributeValue("cy", "635000");
            AssertXmlEquivalent(expected, XElement.Parse(ReadXml(savedSlide.GetPackagePart()).OuterXml));
            Assert.That(RelationshipKeys(savedSlide.GetPackagePart()), Is.EqualTo(originalRelationships));
        }

        private static List<string> RelationshipKeys(PackagePart part)
        {
            var keys = new List<string>();
            foreach (PackageRelationship relationship in part.Relationships)
            {
                keys.Add(relationship.Id + ":" + relationship.RelationshipType);
            }
            return keys;
        }

        [Test]
        public void MissingExplicitTransformsFailExplicitly()
        {
            using var input = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("layouts.pptx");
            using var presentation = new XMLSlideShow(input);
            var layout = presentation.GetSlideMasters()[0].GetSlideLayouts().First(item => item.Type == "title");
            var shapes = layout.GetShapes();
            Assert.That(shapes.Count, Is.GreaterThan(0));
            var withoutTransform = shapes.OfType<XSLFSimpleShape>().First(shape => !HasTransform(shape));
            Assert.That(() => withoutTransform.Bounds, Throws.TypeOf<POIXMLException>()
                .With.Message.Contain("no explicit transform"));
            Assert.That(() => withoutTransform.Bounds = SKRect.Create(1, 2, 3, 4),
                Throws.TypeOf<POIXMLException>().With.Message.Contain("no explicit transform"));
            Assert.That(withoutTransform.ShapeId, Is.GreaterThan(0U));
            Assert.That(withoutTransform.ShapeName, Is.Not.Null);
        }

        [TestCase("nvSpPr")]
        [TestCase("cNvPr")]
        public void MalformedSupportedShapesFailEnumerationBeforeMutation(string removed)
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            var element = shape.GetXmlObject();
            var removedNode = removed == "nvSpPr"
                ? (XmlNode)element["nvSpPr", CT_Presentation.Namespace]
                : element["nvSpPr", CT_Presentation.Namespace]["cNvPr", CT_Presentation.Namespace];
            removedNode.ParentNode.RemoveChild(removedNode);

            Assert.That(() => slide.GetShapes(), Throws.TypeOf<POIXMLException>());
        }

        [Test]
        public void MissingShapePropertiesFailAtBoundsAccessRatherThanEnumeration()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            var element = shape.GetXmlObject();
            element.RemoveChild(element["spPr", CT_Presentation.Namespace]);

            // A shape without properties still enumerates, but its explicit bounds are absent.
            var enumerated = (XSLFSimpleShape)slide.GetShapes()[0];
            Assert.That(enumerated.ShapeId, Is.EqualTo(shape.ShapeId));
            Assert.That(() => enumerated.Bounds, Throws.TypeOf<POIXMLException>());
            Assert.That(() => enumerated.Bounds = SKRect.Create(1, 2, 3, 4),
                Throws.TypeOf<POIXMLException>());
        }

        [Test]
        public void AmbiguousSupportedShapesFailEnumeration()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            var element = shape.GetXmlObject();
            var properties = element["spPr", CT_Presentation.Namespace];
            properties.AppendChild(properties.OwnerDocument.CreateElement(
                "xfrm", XSLFNamespace));
            Assert.That(() => shape.Bounds, Throws.TypeOf<POIXMLException>()
                .With.Message.Contain("at most one xfrm"));
        }

        [TestCase("0")]
        [TestCase("-1")]
        [TestCase("4294967296")]
        [TestCase("invalid")]
        [TestCase("")]
        public void InvalidShapeIdentifiersFailEnumeration(string id)
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            shape.GetXmlObject()["nvSpPr", CT_Presentation.Namespace]["cNvPr", CT_Presentation.Namespace]
                .SetAttribute("id", id);
            Assert.That(() => slide.GetShapes(), Throws.TypeOf<POIXMLException>()
                .With.Message.Contain("Invalid shape identifier"));
        }

        [Test]
        public void NullNamesAndInvalidGeometryAreRejectedWithoutMutation()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            Assert.That(() => slide.CreateSimpleShape(PresetGeometry.Rectangle, null),
                Throws.TypeOf<ArgumentNullException>());
            Assert.That(() => slide.CreateGroup(null), Throws.TypeOf<ArgumentNullException>());
            Assert.That(() => slide.CreateSimpleShape((PresetGeometry)int.MaxValue, "Shape"),
                Throws.TypeOf<ArgumentOutOfRangeException>());
            Assert.That(slide.GetShapes(), Is.Empty);
        }

        [Test]
        public void NamesRejectNullBeforeMutation()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            var group = slide.CreateGroup("Group");
            Assert.That(() => shape.ShapeName = null, Throws.TypeOf<ArgumentNullException>());
            Assert.That(() => group.ShapeName = null, Throws.TypeOf<ArgumentNullException>());
            Assert.That(shape.ShapeName, Is.EqualTo("Rectangle"));
            Assert.That(group.ShapeName, Is.EqualTo("Group"));
        }

        [Test]
        public void ReadOnlySheetsCannotCreateOrEditShapes()
        {
            using var presentation = new XMLSlideShow();
            presentation.CreateSlide();
            using var package = OPCPackage.Open(new MemoryStream(Save(presentation)), true);
            using var readOnly = new XMLSlideShow(package);
            var slide = readOnly.GetSlides()[0];
            var shapes = slide.GetShapes();
            Assert.That(() => slide.CreateSimpleShape(PresetGeometry.Rectangle, "Shape"),
                Throws.TypeOf<InvalidOperationException>());
            Assert.That(() => slide.CreateGroup("Group"), Throws.TypeOf<InvalidOperationException>());
            if (shapes.Count > 0)
            {
                var shape = (XSLFSimpleShape)shapes[0];
                Assert.That(() => shape.ShapeName = "Edited", Throws.TypeOf<InvalidOperationException>());
                Assert.That(() => shape.Bounds = SKRect.Create(1, 2, 3, 4),
                    Throws.TypeOf<InvalidOperationException>());
            }
            Assert.That(slide.GetShapes(), Is.Empty);
        }

        [Test]
        public void DisposedPresentationsCannotCreateOrEditShapes()
        {
            var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            presentation.Dispose();
            Assert.That(() => slide.CreateSimpleShape(PresetGeometry.Rectangle, "Shape"),
                Throws.TypeOf<ObjectDisposedException>());
            Assert.That(() => slide.CreateGroup("Group"), Throws.TypeOf<ObjectDisposedException>());
            Assert.That(() => shape.ShapeName = "Edited", Throws.TypeOf<ObjectDisposedException>());
            Assert.That(() => shape.Bounds = SKRect.Create(1, 2, 3, 4),
                Throws.TypeOf<ObjectDisposedException>());
        }

        [Test]
        public void DetachedShapesAreRejected()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var shape = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            var group = slide.CreateGroup("Group");
            shape.GetXmlObject().ParentNode.RemoveChild(shape.GetXmlObject());
            group.GetXmlObject().ParentNode.RemoveChild(group.GetXmlObject());

            Assert.That(() => shape.ShapeId, Throws.TypeOf<InvalidOperationException>()
                .With.Message.Contain("detached"));
            Assert.That(() => shape.ShapeName, Throws.TypeOf<InvalidOperationException>());
            Assert.That(() => shape.Bounds, Throws.TypeOf<InvalidOperationException>());
            Assert.That(() => shape.ShapeName = "x", Throws.TypeOf<InvalidOperationException>());
            Assert.That(() => group.GetShapes(), Throws.TypeOf<InvalidOperationException>());
            Assert.That(() => group.CreateSimpleShape(PresetGeometry.Rectangle, "Inner"),
                Throws.TypeOf<InvalidOperationException>());
            Assert.That(() => group.InteriorBounds = SKRect.Create(0, 0, 1, 1),
                Throws.TypeOf<InvalidOperationException>());
        }

        [Test]
        public void DetachedGroupChildrenAreNoLongerEnumerated()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var group = slide.CreateGroup("Group");
            var inner = group.CreateSimpleShape(PresetGeometry.Rectangle, "Inner");
            var kept = group.CreateSimpleShape(PresetGeometry.Rectangle, "Kept");
            inner.GetXmlObject().ParentNode.RemoveChild(inner.GetXmlObject());

            Assert.That(() => inner.ShapeId, Throws.TypeOf<InvalidOperationException>());
            Assert.That(group.GetShapes().Select(shape => shape.ShapeId),
                Is.EqualTo(new uint[] { kept.ShapeId }));
        }

        private const string XSLFNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";

        private static IEnumerable<XSLFSheet> Sheets(XMLSlideShow presentation)
        {
            foreach (var slide in presentation.GetSlides())
            {
                yield return slide;
            }
            foreach (var master in presentation.GetSlideMasters())
            {
                yield return master;
                foreach (var layout in master.GetSlideLayouts())
                {
                    yield return layout;
                }
            }
        }

        private static bool HasTransform(XSLFSimpleShape shape)
        {
            return shape.GetXmlObject()["spPr", CT_Presentation.Namespace]?["xfrm", XSLFNamespace] != null;
        }

        private static XmlElement FindTransform(XSLFShape shape)
        {
            var element = shape.GetXmlObject();
            var holderName = element.LocalName == "grpSp" ? "grpSpPr" : "spPr";
            return (XmlElement)element[holderName, CT_Presentation.Namespace]["xfrm", XSLFNamespace];
        }

        private static string GetAttribute(XmlElement transform, string child, string name)
        {
            return transform[child, XSLFNamespace].GetAttribute(name);
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

        private static void AssertXmlEquivalent(XElement expected, XElement actual)
        {
            expected.DescendantsAndSelf().Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
            actual.DescendantsAndSelf().Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
            Assert.That(XNode.DeepEquals(expected, actual), Is.True, "PresentationML content changed.");
        }
    }

    /// <summary>
    /// Exposes the internal XML helper to tests in this assembly.
    /// </summary>
    internal static class XSLFXmlHelperProxy
    {
        internal static XmlElement CreateElement(string xml)
        {
            var document = new XmlDocument { XmlResolver = null };
            document.LoadXml(xml);
            return document.DocumentElement;
        }
    }
}
