using System;
using System.Globalization;
using System.Xml;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.Util;
using NPOI.XDDF.UserModel;
using SkiaSharp;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A preset-geometry shape that supports reading and editing its name and explicit
    /// position and size in points. Rotation, flip transforms, styling, and text are preserved
    /// but cannot be edited in this implementation.
    /// </summary>
    public class XSLFSimpleShape : XSLFShape
    {
        private readonly XmlElement element;

        internal XSLFSimpleShape(XSLFShapeContainer container, XmlElement element)
            : base(container)
        {
            this.element = element ?? throw new ArgumentNullException(nameof(element));
            XSLFShapeTree.ReadShapeId(element, "nvSpPr");
        }

        /// <summary>
        /// Gets or sets the shape name shown by presentation applications.
        /// </summary>
        public string ShapeName
        {
            get
            {
                EnsureAttached();
                return ReadName();
            }
            set
            {
                EnsureEditable();
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value));
                }
                XSLFShapeTree.GetNonVisualProperties(element, "nvSpPr", "cNvPr").SetAttribute("name", value);
            }
        }

        /// <summary>
        /// Gets or sets the explicit position and size of this shape in points, measured in the
        /// coordinate system of the containing tree. Setting bounds rewrites only the offset and
        /// extent of the explicit transform, preserving rotation and flip attributes.
        /// </summary>
        /// <exception cref="POIXMLException">The shape has no explicit transform.</exception>
        /// <exception cref="ArgumentException">The bounds have a negative width or height.</exception>
        public SKRect Bounds
        {
            get
            {
                EnsureAttached();
                return ReadBounds(XSLFShapeTree.GetTransform(element, "spPr"));
            }
            set
            {
                EnsureEditable();
                if (value.Width < 0 || value.Height < 0)
                {
                    throw new ArgumentException("The shape bounds cannot have a negative extent.", nameof(value));
                }
                WriteBounds(XSLFShapeTree.GetTransform(element, "spPr"), value, "off", "ext");
            }
        }

        /// <summary>
        /// Gets the preset geometry of this shape. Custom geometries are reported as opaque
        /// shapes and never wrapped by this class.
        /// </summary>
        public PresetGeometry Geometry
        {
            get
            {
                EnsureAttached();
                var properties = XSLFShapeTree.GetRequiredChild(element, CT_Presentation.Namespace, "spPr");
                var geometry = XSLFShapeTree.GetUniqueChild(properties, XSLFXmlHelper.DrawingNamespace, "prstGeom");
                if (geometry == null)
                {
                    throw new POIXMLException("The shape has no preset geometry.");
                }
                var prst = geometry.GetAttribute("prst");
                foreach (NPOI.OpenXmlFormats.Dml.ST_ShapeType type in
                    Enum.GetValues(typeof(NPOI.OpenXmlFormats.Dml.ST_ShapeType)))
                {
                    if (string.Equals(type.ToString(), prst, StringComparison.Ordinal))
                    {
                        return PresetGeometryExtensions.ValueOf(type);
                    }
                }
                throw new POIXMLException("Unknown preset geometry: " + prst);
            }
        }

        internal override XmlElement GetElement()
        {
            return element;
        }

        internal override uint ReadShapeId()
        {
            return XSLFShapeTree.ReadShapeId(element, "nvSpPr");
        }

        internal static SKRect ReadBounds(XmlElement transform)
        {
            var offset = XSLFShapeTree.GetUniqueChild(transform, XSLFXmlHelper.DrawingNamespace, "off");
            var extent = XSLFShapeTree.GetUniqueChild(transform, XSLFXmlHelper.DrawingNamespace, "ext");
            if (offset == null || extent == null)
            {
                throw new POIXMLException("The transform has no offset or extent.");
            }
            return ReadRect(offset, extent);
        }

        internal static SKRect ReadRect(XmlElement offset, XmlElement extent)
        {
            return SKRect.Create(
                ToPoint(ReadCoordinate(offset, "x")),
                ToPoint(ReadCoordinate(offset, "y")),
                ToPoint(ReadCoordinate(extent, "cx")),
                ToPoint(ReadCoordinate(extent, "cy")));
        }

        internal static void WriteBounds(XmlElement transform, SKRect bounds, string offsetName, string extentName)
        {
            var offset = XSLFShapeTree.GetRequiredChild(transform, XSLFXmlHelper.DrawingNamespace, offsetName);
            var extent = XSLFShapeTree.GetRequiredChild(transform, XSLFXmlHelper.DrawingNamespace, extentName);
            WriteCoordinate(offset, "x", bounds.Left);
            WriteCoordinate(offset, "y", bounds.Top);
            WriteCoordinate(extent, "cx", bounds.Width);
            WriteCoordinate(extent, "cy", bounds.Height);
        }

        private string ReadName()
        {
            return XSLFShapeTree.GetNonVisualProperties(element, "nvSpPr", "cNvPr").GetAttribute("name");
        }

        internal static double ReadCoordinate(XmlElement parent, string name)
        {
            var value = parent.GetAttribute(name);
            if (!long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var emu))
            {
                throw new POIXMLException("Invalid EMU coordinate in " + parent.LocalName + ": " + value);
            }
            return Units.ToPoints(emu);
        }

        private static float ToPoint(double value)
        {
            if (value < float.MinValue || value > float.MaxValue)
            {
                throw new POIXMLException("The coordinate does not fit in points: " + value);
            }
            return (float)value;
        }

        internal static void WriteCoordinate(XmlElement parent, string name, double points)
        {
            var emu = points * Units.EMU_PER_POINT;
            if (emu > int.MaxValue || emu < int.MinValue)
            {
                throw new POIXMLException("The coordinate does not fit in EMUs: " +
                    points.ToString(CultureInfo.InvariantCulture));
            }
            if (emu < 0 && name[0] == 'c')
            {
                throw new ArgumentException("Extents cannot be negative.", nameof(points));
            }
            parent.SetAttribute(name, ((int)Math.Round(emu, MidpointRounding.AwayFromZero))
                .ToString(CultureInfo.InvariantCulture));
        }
    }
}
