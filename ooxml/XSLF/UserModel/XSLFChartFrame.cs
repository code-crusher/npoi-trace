using System;
using System.Globalization;
using System.Xml;
using NPOI.OpenXmlFormats.Presentation;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A PresentationML graphic frame that displays an existing chart in a sheet's shape tree.
    /// The frame name and explicit point-based anchor can be edited; chart contents are not modified.
    /// </summary>
    public class XSLFChartFrame
    {
        private readonly XSLFSheet sheet;
        private readonly XmlElement frame;

        internal XSLFChartFrame(XSLFSheet sheet, XmlElement frame)
        {
            this.sheet = sheet;
            this.frame = frame;
        }

        /// <summary>
        /// Gets the sheet containing this frame.
        /// </summary>
        public XSLFSheet GetSheet()
        {
            return sheet;
        }

        /// <summary>
        /// Gets or sets the frame name shown by presentation applications.
        /// </summary>
        /// <exception cref="ArgumentException">The name is null or empty.</exception>
        public string Name
        {
            get
            {
                var property = NonVisualProperty;
                return property?.GetAttribute("name") ?? string.Empty;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new ArgumentException("The chart frame name must not be null or empty.", nameof(value));
                }
                var property = NonVisualProperty;
                if (property == null)
                {
                    throw new POIXMLException("The chart frame has no non-visual drawing properties: " +
                        sheet.GetPackagePart().PartName.Name);
                }
                property.SetAttribute("name", value);
            }
        }

        /// <summary>
        /// Gets or sets the explicit point-based frame anchor in EMUs.
        /// The setter requires an explicit offset and extent, as produced by this API.
        /// </summary>
        /// <exception cref="POIXMLException">The frame has no or a non-explicit transform.</exception>
        public XSLFChartAnchor Anchor
        {
            get
            {
                var offset = Offset;
                var extent = Extent;
                return new XSLFChartAnchor(
                    ParseEmu(offset, "x", "xfrm off"),
                    ParseEmu(offset, "y", "xfrm off"),
                    ParseEmu(extent, "cx", "xfrm ext"),
                    ParseEmu(extent, "cy", "xfrm ext"));
            }
            set
            {
                var offset = Offset;
                var extent = Extent;
                if (offset == null || extent == null)
                {
                    throw new POIXMLException("Only explicit point-based chart frame anchors are supported: " +
                        sheet.GetPackagePart().PartName.Name);
                }
                offset.SetAttribute("x", value.X.ToString(CultureInfo.InvariantCulture));
                offset.SetAttribute("y", value.Y.ToString(CultureInfo.InvariantCulture));
                extent.SetAttribute("cx", value.Width.ToString(CultureInfo.InvariantCulture));
                extent.SetAttribute("cy", value.Height.ToString(CultureInfo.InvariantCulture));
            }
        }

        private XmlElement NonVisualProperty =>
            (XmlElement)frame.SelectSingleNode(
                "p:nvGraphicFramePr/p:cNvPr", NamespaceManager) ?? (XmlElement)frame.SelectSingleNode(
                "nvGraphicFramePr/cNvPr", NamespaceManager);

        private XmlElement Offset =>
            (XmlElement)frame.SelectSingleNode("p:xfrm/a:off", NamespaceManager);

        private XmlElement Extent =>
            (XmlElement)frame.SelectSingleNode("p:xfrm/a:ext", NamespaceManager);

        private XmlNamespaceManager NamespaceManager
        {
            get
            {
                var manager = new XmlNamespaceManager(frame.OwnerDocument.NameTable);
                manager.AddNamespace("p", CT_Presentation.Namespace);
                manager.AddNamespace("a", XSLFXmlHelper.DrawingNamespace);
                return manager;
            }
        }

        private long ParseEmu(XmlElement element, string attribute, string location)
        {
            if (element == null)
            {
                throw new POIXMLException("The chart frame has no " + location + " element: " +
                    sheet.GetPackagePart().PartName.Name);
            }
            var value = element.GetAttribute(attribute);
            if (!long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var emu))
            {
                throw new POIXMLException("Invalid chart frame anchor value '" + value + "' in " +
                    sheet.GetPackagePart().PartName.Name);
            }
            return emu;
        }
    }
}
