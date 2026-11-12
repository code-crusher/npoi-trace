using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using NPOI.OpenXmlFormats.Presentation;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// Locates and edits the shape tree of a common slide data element.
    /// </summary>
    internal static class XSLFShapeTree
    {
        internal static XmlElement GetShapeTree(XmlElement cSld)
        {
            if (cSld == null)
            {
                throw new POIXMLException("The sheet has no common slide data.");
            }
            var candidates = cSld.GetElementsByTagName("spTree", CT_Presentation.Namespace);
            if (candidates.Count != 1)
            {
                throw new POIXMLException("The common slide data must contain exactly one shape tree: " +
                    candidates.Count);
            }
            return (XmlElement)candidates[0];
        }

        internal static XmlElement GetNonVisualProperties(XmlElement shape, string listName, string propertyName)
        {
            var list = GetSingleChild(shape, CT_Presentation.Namespace, listName);
            return GetSingleChild(list, CT_Presentation.Namespace, propertyName);
        }

        internal static XmlElement GetUniqueChild(XmlElement parent, string namespaceUri, string localName)
        {
            var children = FindDirectChildren(parent, namespaceUri, localName);
            if (children.Count > 1)
            {
                throw new POIXMLException("A " + parent.LocalName + " must contain at most one " + localName + ": " +
                    children.Count);
            }
            return children.Count == 0 ? null : children[0];
        }

        internal static XmlElement GetSingleChild(XmlElement parent, string namespaceUri, string localName)
        {
            var children = FindDirectChildren(parent, namespaceUri, localName);
            if (children.Count != 1)
            {
                throw new POIXMLException("A " + parent.LocalName + " must contain exactly one " + localName + ": " +
                    children.Count);
            }
            return children[0];
        }

        internal static XmlElement GetRequiredChild(XmlElement parent, string namespaceUri, string localName)
        {
            var child = GetUniqueChild(parent, namespaceUri, localName);
            if (child == null)
            {
                throw new POIXMLException("The " + parent.LocalName + " has no " + localName + ".");
            }
            return child;
        }

        private static List<XmlElement> FindDirectChildren(XmlElement parent, string namespaceUri, string localName)
        {
            var children = new List<XmlElement>();
            foreach (XmlNode node in parent.ChildNodes)
            {
                if (node is XmlElement element && element.NamespaceURI == namespaceUri &&
                    element.LocalName == localName)
                {
                    children.Add(element);
                }
            }
            return children;
        }

        internal static uint ReadShapeId(XmlElement shape, string listName)
        {
            var property = GetNonVisualProperties(shape, listName, "cNvPr");
            var value = property.GetAttribute("id");
            if (!uint.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) ||
                id < 1 || id > int.MaxValue)
            {
                throw new POIXMLException("Invalid shape identifier in a " + shape.LocalName + ": " + value);
            }
            return id;
        }

        internal static XmlElement GetTransform(XmlElement shape, string holderName)
        {
            var holder = GetSingleChild(shape, CT_Presentation.Namespace, holderName);
            var transform = GetUniqueChild(holder, XSLFXmlHelper.DrawingNamespace, "xfrm");
            if (transform == null)
            {
                throw new POIXMLException("A " + shape.LocalName + " has no explicit transform.");
            }
            return transform;
        }

        internal static void CollectUsedShapeIds(XmlElement root, HashSet<uint> usedIds)
        {
            foreach (XmlElement property in root.GetElementsByTagName("cNvPr", CT_Presentation.Namespace))
            {
                uint.TryParse(property.GetAttribute("id"), NumberStyles.Integer, CultureInfo.InvariantCulture,
                    out var id);
                usedIds.Add(id);
            }
        }

        internal static uint AllocateShapeId(IEnumerable<XmlElement> sheetRoots, XmlElement tree)
        {
            var usedIds = new HashSet<uint>();
            foreach (var root in sheetRoots)
            {
                if (root != null)
                {
                    CollectUsedShapeIds(root, usedIds);
                }
            }
            CollectUsedShapeIds(tree, usedIds);
            for (var id = 2U; id <= int.MaxValue; id++)
            {
                if (usedIds.Add(id))
                {
                    return id;
                }
            }
            throw new InvalidOperationException("No more shape identifiers are available.");
        }
    }
}
