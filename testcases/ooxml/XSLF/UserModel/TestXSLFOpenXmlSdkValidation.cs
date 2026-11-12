using System;
using System.IO;
using System.Linq;
using NPOI;
using NPOI.XSLF.UserModel;
using NPOI.XDDF.UserModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using NUnit.Framework;
using SkiaSharp;

namespace TestCases.XSLF.UserModel
{
    /// <summary>
    /// Validates packages produced by the NPOI shape container APIs with the independent
    /// Open XML SDK implementation and schema validator.
    /// </summary>
    [TestFixture]
    public class TestXSLFOpenXmlSdkValidation
    {
        [Test]
        public void CreatedShapesAndGroupsProduceSchemaValidPackages()
        {
            using var presentation = new XMLSlideShow();
            var slide = presentation.CreateSlide();
            var rectangle = slide.CreateSimpleShape(PresetGeometry.Rectangle, "Rectangle");
            rectangle.Bounds = SKRect.Create(72, 36, 144, 72);
            var group = slide.CreateGroup("Group");
            group.Bounds = SKRect.Create(10, 20, 200, 100);
            group.InteriorBounds = SKRect.Create(0, 0, 200, 100);
            var ellipse = group.CreateSimpleShape(PresetGeometry.Ellipse, "Ellipse");
            ellipse.Bounds = SKRect.Create(5, 5, 50, 25);
            var nested = group.CreateGroup("Nested");
            nested.Bounds = SKRect.Create(1, 2, 10, 5);
            slide.CreateGroup("Empty");

            using var package = OpenPackage(presentation);
            AssertNoSchemaErrors(package);
            var slidePart = package.PresentationPart.SlideParts.Single();
            var tree = slidePart.Slide.CommonSlideData.ShapeTree;
            Assert.That(tree, Is.Not.Null);
            var shapes = tree.Elements<DocumentFormat.OpenXml.Presentation.Shape>().ToArray();
            Assert.That(shapes.Length, Is.EqualTo(1), "The slide holds one simple shape.");
            Assert.That(shapes[0].NonVisualShapeProperties.NonVisualDrawingProperties.Id.Value,
                Is.EqualTo(2U));
            Assert.That(shapes[0].ShapeProperties.Transform2D.Offset.X.Value, Is.EqualTo(914400L));
            Assert.That(shapes[0].ShapeProperties.Transform2D.Extents.Cx.Value, Is.EqualTo(1828800L));
            var groups = tree.Elements<DocumentFormat.OpenXml.Presentation.GroupShape>().ToArray();
            Assert.That(groups.Length, Is.EqualTo(2), "The slide holds two groups.");
            var edited = groups[0];
            Assert.That(edited.NonVisualGroupShapeProperties.NonVisualDrawingProperties.Id.Value,
                Is.EqualTo(3U));
            Assert.That(edited.NonVisualGroupShapeProperties.NonVisualDrawingProperties.Name.Value,
                Is.EqualTo("Group"));
            var transform = (DocumentFormat.OpenXml.Drawing.TransformGroup)
                edited.GroupShapeProperties.TransformGroup;
            Assert.That(transform.Offset.X.Value, Is.EqualTo(127000L));
            Assert.That(transform.Offset.Y.Value, Is.EqualTo(254000L));
            Assert.That(transform.Extents.Cx.Value, Is.EqualTo(2540000L));
            Assert.That(transform.Extents.Cy.Value, Is.EqualTo(1270000L));
            Assert.That(transform.ChildOffset.X.Value, Is.EqualTo(0L));
            Assert.That(transform.ChildExtents.Cx.Value, Is.EqualTo(2540000L));
            var nestedGroup = edited.Elements<DocumentFormat.OpenXml.Presentation.GroupShape>().Single();
            Assert.That(nestedGroup.NonVisualGroupShapeProperties.NonVisualDrawingProperties.Name.Value,
                Is.EqualTo("Nested"));
            var nestedShape = edited.Elements<DocumentFormat.OpenXml.Presentation.Shape>().Single();
            Assert.That(nestedShape.NonVisualShapeProperties.NonVisualDrawingProperties.Name.Value,
                Is.EqualTo("Ellipse"));
            Assert.That(nestedShape.ShapeProperties.Transform2D.Offset.X.Value, Is.EqualTo(63500L));
            var geometry = nestedShape.ShapeProperties
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.PresetGeometry>();
            Assert.That(geometry, Is.Not.Null);
            Assert.That(geometry.Preset.Value,
                Is.EqualTo(DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse));
        }

        [Test]
        public void EditingExistingPresentationsStaysSchemaValid()
        {
            using var stream = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("shapes.pptx");
            using var presentation = new XMLSlideShow(stream);
            foreach (var slide in presentation.GetSlides())
            {
                foreach (var shape in slide.GetShapes().OfType<XSLFSimpleShape>())
                {
                    shape.ShapeName = shape.ShapeName + " Edited";
                    try
                    {
                        shape.Bounds = SKRect.Create(10, 10, 100, 100);
                    }
                    catch (POIXMLException)
                    {
                        // Placeholders without an explicit transform keep their inherited
                        // position; the exception path is covered by the main test suite.
                    }
                }
                var added = slide.CreateSimpleShape(PresetGeometry.RightArrow, "Added Arrow");
                added.Bounds = SKRect.Create(50, 50, 120, 60);
                var group = slide.CreateGroup("Added Group");
                group.Bounds = SKRect.Create(5, 5, 60, 30);
                group.InteriorBounds = SKRect.Create(0, 0, 120, 60);
                var inner = group.CreateSimpleShape(PresetGeometry.Rectangle, "Inner");
                inner.Bounds = SKRect.Create(0, 0, 60, 30);
            }

            using var package = OpenPackage(presentation);
            AssertNoSchemaErrors(package);
            Assert.That(package.PresentationPart.SlideParts.Count(),
                Is.EqualTo(presentation.GetSlides().Count));
        }

        [Test]
        public void EditingShapesOnLayoutAndMasterStaysSchemaValid()
        {
            using var presentation = new XMLSlideShow();
            var master = presentation.GetSlideMasters()[0];
            var layout = master.GetSlideLayouts()[0];
            var masterShape = master.CreateSimpleShape(PresetGeometry.Rectangle, "Master Shape");
            masterShape.Bounds = SKRect.Create(0, 0, 100, 100);
            var layoutShape = layout.CreateSimpleShape(PresetGeometry.Ellipse, "Layout Shape");
            layoutShape.Bounds = SKRect.Create(-36, -18, 72, 54);
            var layoutGroup = layout.CreateGroup("Layout Group");
            layoutGroup.Bounds = SKRect.Create(0, 0, 10, 10);
            layoutGroup.InteriorBounds = SKRect.Create(0, 0, 10, 10);

            using var package = OpenPackage(presentation);
            AssertNoSchemaErrors(package);
            var layoutPart = package.PresentationPart.SlideMasterParts.Single()
                .SlideLayoutParts.Single();
            var groups = layoutPart.SlideLayout.CommonSlideData.ShapeTree
                .Elements<DocumentFormat.OpenXml.Presentation.GroupShape>().ToArray();
            Assert.That(groups.Single().NonVisualGroupShapeProperties.NonVisualDrawingProperties
                .Name.Value, Is.EqualTo("Layout Group"));
        }

        private static PresentationDocument OpenPackage(XMLSlideShow presentation)
        {
            using var output = new MemoryStream();
            presentation.Write(output);
            // XMLSlideShow.Write closes the stream, so the validator needs a fresh copy.
            return PresentationDocument.Open(new MemoryStream(output.ToArray()), false);
        }

        private static void AssertNoSchemaErrors(PresentationDocument document)
        {
            var errors = new OpenXmlValidator().Validate(document).ToList();
            foreach (var error in errors)
            {
                TestContext.Error.WriteLine(error.Description + " at " + error.Path?.XPath);
            }
            Assert.That(errors, Is.Empty, "The Open XML SDK validator reported schema errors.");
        }
    }
}
