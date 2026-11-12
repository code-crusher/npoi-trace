using System;
using System.Xml;
using NPOI.OpenXmlFormats.Presentation;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A shape in a slide, layout, master, or group shape tree. Unsupported drawing content
    /// is wrapped by <see cref="XSLFOpaqueShape"/> and cannot be edited.
    /// </summary>
    public abstract class XSLFShape
    {
        private readonly XSLFShapeContainer container;

        internal XSLFShape(XSLFShapeContainer container)
        {
            this.container = container ?? throw new ArgumentNullException(nameof(container));
        }

        /// <summary>
        /// Gets the stable shape identifier used by PresentationML references.
        /// </summary>
        /// <exception cref="POIXMLException">The shape has no valid identifier.</exception>
        public virtual uint ShapeId
        {
            get
            {
                EnsureAttached();
                return ReadShapeId();
            }
        }

        /// <summary>
        /// Returns the container that enumerated or created this shape.
        /// </summary>
        public XSLFShapeContainer GetContainer()
        {
            return container;
        }

        /// <summary>
        /// Returns the underlying shape element with its unknown content intact.
        /// </summary>
        public XmlElement GetXmlObject()
        {
            EnsureAttached();
            return GetElement();
        }

        internal abstract XmlElement GetElement();

        internal abstract uint ReadShapeId();

        internal void EnsureAttached()
        {
            var tree = container.GetShapeTreeForValidation();
            for (var node = (XmlNode)GetElement(); node != null; node = node.ParentNode)
            {
                if (ReferenceEquals(node, tree))
                {
                    return;
                }
            }
            throw new InvalidOperationException("The shape is detached from its container.");
        }

        internal void EnsureEditable()
        {
            EnsureAttached();
            container.EnsureCanEdit();
        }
    }
}
