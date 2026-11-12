using System;
using System.Xml;
using NPOI.OpenXmlFormats.Presentation;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// Drawing content that is preserved without interpretation, such as pictures, tables,
    /// graphic frames, connectors, and unknown elements. Only the identifier and the raw XML
    /// are exposed; editing this content is a separate stage of the XSLF roadmap.
    /// </summary>
    public class XSLFOpaqueShape : XSLFShape
    {
        private readonly XmlElement element;

        internal XSLFOpaqueShape(XSLFShapeContainer container, XmlElement element)
            : base(container)
        {
            this.element = element ?? throw new ArgumentNullException(nameof(element));
        }

        /// <summary>
        /// Gets the stable shape identifier of this drawing content, or zero when the element
        /// carries no PresentationML non-visual properties.
        /// </summary>
        public override uint ShapeId
        {
            get
            {
                EnsureAttached();
                return FindIdentifier();
            }
        }

        /// <summary>
        /// Gets the drawing name of this content, or an empty string when the element has none.
        /// </summary>
        public string ShapeName
        {
            get
            {
                EnsureAttached();
                var property = FindNonVisualProperties();
                return property?.GetAttribute("name") ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets the local name of the wrapped drawing element, such as "pic" or "graphicFrame".
        /// </summary>
        public string DrawingType
        {
            get
            {
                EnsureAttached();
                return element.LocalName;
            }
        }

        internal override XmlElement GetElement()
        {
            return element;
        }

        internal override uint ReadShapeId()
        {
            return FindIdentifier();
        }

        private uint FindIdentifier()
        {
            var property = FindNonVisualProperties();
            if (property == null)
            {
                return 0;
            }
            var value = property.GetAttribute("id");
            if (!uint.TryParse(value, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var id) || id == 0)
            {
                throw new POIXMLException("Invalid shape identifier in a " + element.LocalName + ": " + value);
            }
            return id;
        }

        private XmlElement FindNonVisualProperties()
        {
            // Picture, connector, and graphic frame content carries its non-visual properties
            // in a direct child element such as nvPicPr, nvCxnSpPr, or nvGraphicFramePr.
            foreach (XmlNode node in element.ChildNodes)
            {
                if (!(node is XmlElement container) || container.NamespaceURI != CT_Presentation.Namespace ||
                    container.LocalName.Length < 3 || !container.LocalName.EndsWith("Pr", StringComparison.Ordinal))
                {
                    continue;
                }
                foreach (XmlNode child in container.ChildNodes)
                {
                    if (child is XmlElement property && property.LocalName == "cNvPr" &&
                        property.NamespaceURI == CT_Presentation.Namespace)
                    {
                        return property;
                    }
                }
            }
            return null;
        }
    }
}
