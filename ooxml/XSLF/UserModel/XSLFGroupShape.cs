using System;
using System.Collections.Generic;
using System.Xml;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.XDDF.UserModel;
using SkiaSharp;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A group shape that supports reading and editing its name, its explicit position and size in
    /// points, and the interior coordinate system of its children. Children are enumerated and
    /// created through the group itself; their coordinates are measured in the interior system.
    /// Rotation and flip transforms and unknown content are preserved when editing.
    /// </summary>
    public class XSLFGroupShape : XSLFShape, XSLFShapeContainer
    {
        private readonly XmlElement element;

        internal XSLFGroupShape(XSLFShapeContainer container, XmlElement element)
            : base(container)
        {
            this.element = element ?? throw new ArgumentNullException(nameof(element));
            XSLFShapeTree.ReadShapeId(element, "nvGrpSpPr");
        }

        /// <summary>
        /// Gets or sets the group name shown by presentation applications.
        /// </summary>
        public string ShapeName
        {
            get
            {
                EnsureAttached();
                return XSLFShapeTree.GetNonVisualProperties(element, "nvGrpSpPr", "cNvPr").GetAttribute("name");
            }
            set
            {
                EnsureEditable();
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value));
                }
                XSLFShapeTree.GetNonVisualProperties(element, "nvGrpSpPr", "cNvPr").SetAttribute("name", value);
            }
        }

        /// <summary>
        /// Gets or sets the explicit exterior position and size of this group in points, measured
        /// in the coordinate system of the containing tree. The interior coordinate system is not
        /// rescaled by this property; use <see cref="InteriorBounds"/> to edit it.
        /// </summary>
        /// <exception cref="POIXMLException">The group has no explicit transform.</exception>
        /// <exception cref="ArgumentException">The bounds have a negative width or height.</exception>
        public SKRect Bounds
        {
            get
            {
                EnsureAttached();
                return XSLFSimpleShape.ReadBounds(GetRequiredTransform());
            }
            set
            {
                EnsureEditable();
                if (value.Width < 0 || value.Height < 0)
                {
                    throw new ArgumentException("The group bounds cannot have a negative extent.", nameof(value));
                }
                XSLFSimpleShape.WriteBounds(GetRequiredTransform(), value, "off", "ext");
            }
        }

        /// <summary>
        /// Gets or sets the interior coordinate system of this group in points: the offset and
        /// extent of the child coordinate space in which the bounds of child shapes are measured.
        /// </summary>
        /// <exception cref="POIXMLException">The group has no explicit transform.</exception>
        /// <exception cref="ArgumentException">The bounds have a negative width or height.</exception>
        public SKRect InteriorBounds
        {
            get
            {
                EnsureAttached();
                return ReadInteriorBounds(GetRequiredTransform());
            }
            set
            {
                EnsureEditable();
                if (value.Width < 0 || value.Height < 0)
                {
                    throw new ArgumentException("The interior bounds cannot have a negative extent.", nameof(value));
                }
                WriteInteriorBounds(GetRequiredTransform(), value);
            }
        }

        /// <inheritdoc/>
        public IReadOnlyList<XSLFShape> GetShapes()
        {
            EnsureAttached();
            return XSLFShapeContainerSupport.GetShapes(this, element);
        }

        /// <inheritdoc/>
        public XSLFSimpleShape CreateSimpleShape(PresetGeometry geometry, string name)
        {
            return XSLFShapeContainerSupport.CreateSimpleShape(this, element, GetSheetRoots(),
                geometry, name);
        }

        /// <inheritdoc/>
        public XSLFGroupShape CreateGroup(string name)
        {
            return XSLFShapeContainerSupport.CreateGroup(this, element, GetSheetRoots(), name);
        }

        internal override XmlElement GetElement()
        {
            return element;
        }

        internal override uint ReadShapeId()
        {
            return XSLFShapeTree.ReadShapeId(element, "nvGrpSpPr");
        }

        internal XmlElement GetRequiredTransform()
        {
            var transform = XSLFShapeTree.GetTransform(element, "grpSpPr");
            if (transform == null)
            {
                throw new POIXMLException("The group has no explicit transform.");
            }
            return transform;
        }

        XmlElement XSLFShapeContainer.GetShapeTreeForValidation()
        {
            return element;
        }

        IEnumerable<XmlElement> XSLFShapeContainer.GetSheetElementForValidation()
        {
            // Groups are nested inside the sheet part, so identifiers are allocated against the
            // whole sheet to stay unique across every tree and opaque branch.
            return GetContainer().GetSheetElementForValidation();
        }

        void XSLFShapeContainer.EnsureCanEdit()
        {
            EnsureEditable();
        }

        private IEnumerable<XmlElement> GetSheetRoots()
        {
            return GetContainer().GetSheetElementForValidation();
        }

        private static SKRect ReadInteriorBounds(XmlElement transform)
        {
            var offset = XSLFShapeTree.GetRequiredChild(transform, XSLFXmlHelper.DrawingNamespace, "chOff");
            var extent = XSLFShapeTree.GetRequiredChild(transform, XSLFXmlHelper.DrawingNamespace, "chExt");
            return XSLFSimpleShape.ReadRect(offset, extent);
        }

        private static void WriteInteriorBounds(XmlElement transform, SKRect bounds)
        {
            var offset = XSLFShapeTree.GetRequiredChild(transform, XSLFXmlHelper.DrawingNamespace, "chOff");
            var extent = XSLFShapeTree.GetRequiredChild(transform, XSLFXmlHelper.DrawingNamespace, "chExt");
            XSLFSimpleShape.WriteCoordinate(offset, "x", bounds.Left);
            XSLFSimpleShape.WriteCoordinate(offset, "y", bounds.Top);
            XSLFSimpleShape.WriteCoordinate(extent, "cx", bounds.Width);
            XSLFSimpleShape.WriteCoordinate(extent, "cy", bounds.Height);
        }
    }
}
