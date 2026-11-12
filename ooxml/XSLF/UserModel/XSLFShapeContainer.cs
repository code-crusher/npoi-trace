using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Xml;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.XDDF.UserModel;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A slide, layout, master, or group shape that owns a tree of shapes in document order.
    /// Coordinates are expressed in points and converted to EMUs using the NPOI units.
    /// </summary>
    public interface XSLFShapeContainer
    {
        /// <summary>
        /// Returns the shape children of this container in XML order. Simple and group shapes are
        /// wrapped for editing; other drawing content is returned as <see cref="XSLFOpaqueShape"/>
        /// wrappers that only expose their identifiers and raw XML.
        /// </summary>
        /// <exception cref="POIXMLException">A supported shape has malformed PresentationML.</exception>
        IReadOnlyList<XSLFShape> GetShapes();

        /// <summary>
        /// Appends a new preset shape to this container and returns it.
        /// </summary>
        /// <param name="geometry">The preset geometry of the new shape.</param>
        /// <param name="name">The shape name, which must not be null.</param>
        /// <returns>The appended simple shape.</returns>
        /// <exception cref="ArgumentException">The geometry is not a DrawingML shape preset.</exception>
        /// <exception cref="ArgumentNullException">The name is null.</exception>
        /// <exception cref="InvalidOperationException">
        /// The sheet is read-only, disposed, or no identifier is available.
        /// </exception>
        XSLFSimpleShape CreateSimpleShape(PresetGeometry geometry, string name);

        /// <summary>
        /// Appends a new empty group shape to this container and returns it.
        /// </summary>
        /// <param name="name">The group name, which must not be null.</param>
        /// <returns>The appended group shape.</returns>
        /// <exception cref="ArgumentNullException">The name is null.</exception>
        /// <exception cref="InvalidOperationException">
        /// The sheet is read-only, disposed, or no identifier is available.
        /// </exception>
        XSLFGroupShape CreateGroup(string name);

        /// <summary>
        /// Returns the tree element that holds the shapes of this container.
        /// </summary>
        XmlElement GetShapeTreeForValidation();

        /// <summary>
        /// Returns the root elements of the sheet part that owns this container, used to keep
        /// newly allocated identifiers unique across the whole sheet.
        /// </summary>
        IEnumerable<XmlElement> GetSheetElementForValidation();

        /// <summary>
        /// Rejects mutation when the sheet is read-only, disposed, or detached.
        /// </summary>
        void EnsureCanEdit();
    }

    /// <summary>
    /// Shared shape enumeration and creation for sheets and group shapes.
    /// </summary>
    internal static class XSLFShapeContainerSupport
    {
        internal static ReadOnlyCollection<XSLFShape> GetShapes(XSLFShapeContainer container, XmlElement tree)
        {
            var shapes = new List<XSLFShape>();
            foreach (XmlNode node in tree.ChildNodes)
            {
                if (!(node is XmlElement element))
                {
                    continue;
                }
                // The non-visual properties, group properties, and extension lists of the tree
                // itself are container metadata rather than shapes.
                if (element.NamespaceURI == CT_Presentation.Namespace &&
                    (element.LocalName == "nvGrpSpPr" || element.LocalName == "grpSpPr" ||
                     element.LocalName == "extLst"))
                {
                    continue;
                }
                if (element.NamespaceURI == CT_Presentation.Namespace && element.LocalName == "sp")
                {
                    shapes.Add(new XSLFSimpleShape(container, element));
                }
                else if (element.NamespaceURI == CT_Presentation.Namespace && element.LocalName == "grpSp")
                {
                    shapes.Add(new XSLFGroupShape(container, element));
                }
                else
                {
                    shapes.Add(new XSLFOpaqueShape(container, element));
                }
            }
            return shapes.AsReadOnly();
        }

        internal static XSLFSimpleShape CreateSimpleShape(
            XSLFShapeContainer container, XmlElement tree, IEnumerable<XmlElement> sheetRoots,
            PresetGeometry geometry, string name)
        {
            container.EnsureCanEdit();
            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }
            var type = geometry.ToST_ShapeType();
            var identifier = XSLFShapeTree.AllocateShapeId(sheetRoots, tree);
            var document = tree.OwnerDocument;
            var shape = (XmlElement)document.ImportNode(CreateElement(
                "<p:sp xmlns:p=\"" + CT_Presentation.Namespace + "\" xmlns:a=\"" + XSLFXmlHelper.DrawingNamespace +
                "\"><p:nvSpPr><p:cNvPr id=\"" + Invariant(identifier) + "\" name=\"" + Escape(name) +
                "\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>" +
                "<a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom prst=\"" + type.ToString() +
                "\"><a:avLst/></a:prstGeom></p:spPr></p:sp>"), true);
            tree.AppendChild(shape);
            return new XSLFSimpleShape(container, shape);
        }

        internal static XSLFGroupShape CreateGroup(
            XSLFShapeContainer container, XmlElement tree, IEnumerable<XmlElement> sheetRoots, string name)
        {
            container.EnsureCanEdit();
            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }
            var identifier = XSLFShapeTree.AllocateShapeId(sheetRoots, tree);
            var document = tree.OwnerDocument;
            var group = (XmlElement)document.ImportNode(CreateElement(
                "<p:grpSp xmlns:p=\"" + CT_Presentation.Namespace + "\" xmlns:a=\"" + XSLFXmlHelper.DrawingNamespace +
                "\"><p:nvGrpSpPr><p:cNvPr id=\"" + Invariant(identifier) + "\" name=\"" + Escape(name) +
                "\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/>" +
                "<a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/>" +
                "</a:xfrm></p:grpSpPr></p:grpSp>"), true);
            tree.AppendChild(group);
            return new XSLFGroupShape(container, group);
        }

        private static XmlElement CreateElement(string xml)
        {
            return XSLFXmlHelper.CreateElement(xml);
        }

        private static string Invariant(uint value)
        {
            return value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string Escape(string name)
        {
            return name.Replace("&", "&amp;").Replace("<", "&lt;").Replace("\"", "&quot;");
        }
    }
}
