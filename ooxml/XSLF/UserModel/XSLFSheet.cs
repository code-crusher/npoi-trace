using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;
using NPOI.XDDF.UserModel;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// Shared package navigation, XML persistence, and shape container support for slides,
    /// layouts, and masters.
    /// </summary>
    public abstract class XSLFSheet : POIXMLDocumentPart, XSLFShapeContainer
    {
        private byte[] savedXml;

        protected XSLFSheet()
        {
        }

        protected XSLFSheet(POIXMLDocumentPart parent, PackagePart part)
            : base(parent, part)
        {
        }

        /// <summary>
        /// Returns the presentation containing this sheet.
        /// </summary>
        public XMLSlideShow GetSlideShow()
        {
            var parent = GetParent();
            while (parent != null)
            {
                if (parent is XMLSlideShow slideShow)
                {
                    return slideShow;
                }
                parent = parent.GetParent();
            }
            throw new POIXMLException("The sheet is not attached to a presentation.");
        }

        /// <summary>
        /// Returns the layout or master used by this sheet, or null for a slide master.
        /// </summary>
        public abstract XSLFSheet GetMasterSheet();

        /// <summary>
        /// Returns the inherited theme, or null if the master has no theme.
        /// </summary>
        public abstract XSLFTheme GetTheme();

        /// <inheritdoc/>
        public virtual IReadOnlyList<XSLFShape> GetShapes()
        {
            return XSLFShapeContainerSupport.GetShapes(this, GetShapeTree());
        }

        /// <inheritdoc/>
        public XSLFSimpleShape CreateSimpleShape(PresetGeometry geometry, string name)
        {
            return XSLFShapeContainerSupport.CreateSimpleShape(this, GetShapeTree(), GetSheetRoots(),
                geometry, name);
        }

        /// <inheritdoc/>
        public XSLFGroupShape CreateGroup(string name)
        {
            return XSLFShapeContainerSupport.CreateGroup(this, GetShapeTree(), GetSheetRoots(), name);
        }

        protected void InitializeXml()
        {
            using var output = new MemoryStream();
            WriteXml(output);
            savedXml = output.ToArray();
        }

        protected abstract void WriteXml(Stream stream);

        protected internal override void PrepareForCommit()
        {
            // Keep the original part bytes until serialization confirms that the model changed.
        }

        protected internal override void Commit()
        {
            using var output = new MemoryStream();
            WriteXml(output);
            var xml = output.ToArray();
            if (savedXml != null && savedXml.SequenceEqual(xml))
            {
                return;
            }
            var part = GetPackagePart();
            part.Clear();
            using var stream = part.GetOutputStream();
            stream.Write(xml, 0, xml.Length);
            savedXml = xml;
        }

        internal T GetRequiredRelation<T>(XSLFRelation descriptor) where T : POIXMLDocumentPart
        {
            var relationships = GetPackagePart().GetRelationshipsByType(descriptor.Relation);
            if (relationships.Size != 1)
            {
                throw new POIXMLException("Expected one " + descriptor.Relation + " relationship in " +
                    GetPackagePart().PartName.Name);
            }
            return XSLFXmlHelper.GetReferencedPart<T>(this, relationships.GetRelationship(0).Id, descriptor);
        }

        XmlElement XSLFShapeContainer.GetShapeTreeForValidation()
        {
            return GetShapeTree();
        }

        IEnumerable<XmlElement> XSLFShapeContainer.GetSheetElementForValidation()
        {
            return GetSheetRoots();
        }

        void XSLFShapeContainer.EnsureCanEdit()
        {
            EnsureCanEdit();
        }

        internal abstract IEnumerable<XmlElement> GetSheetRoots();

        internal abstract XmlElement GetShapeTree();

        internal void EnsureCanEdit()
        {
            var package = GetSlideShow().Package;
#if NET8_0_OR_GREATER
            ObjectDisposedException.ThrowIf(package == null, this);
#else
            if (package == null)
            {
                throw new ObjectDisposedException(GetType().Name);
            }
#endif
            if (package.GetPackageAccess() == PackageAccess.READ)
            {
                throw new InvalidOperationException("Cannot edit shapes on a read-only sheet.");
            }
        }
    }
}
