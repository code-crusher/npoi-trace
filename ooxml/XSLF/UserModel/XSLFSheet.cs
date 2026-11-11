using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// Shared package navigation and XML persistence for slides, layouts, and masters.
    /// </summary>
    public abstract class XSLFSheet : POIXMLDocumentPart
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

        /// <summary>
        /// Returns the chart frames in this sheet's shape tree, in document order.
        /// Frames inside nested group shapes are included.
        /// </summary>
        public IReadOnlyList<XSLFChartFrame> GetChartFrames()
        {
            var frames = new List<XSLFChartFrame>();
            var tree = GetShapeTree();
            if (tree == null)
            {
                return frames.AsReadOnly();
            }
            CollectChartFrames(tree, frames);
            return frames.AsReadOnly();
        }

        /// <summary>
        /// Reuses an existing chart in this presentation by adding a chart frame to this sheet.
        /// The chart part itself is not modified or copied.
        /// </summary>
        /// <param name="chart">A chart in this presentation.</param>
        /// <returns>The new chart frame.</returns>
        /// <exception cref="ArgumentNullException">The chart is null.</exception>
        /// <exception cref="ArgumentException">The chart belongs to another presentation or package.</exception>
        /// <exception cref="POIXMLException">The chart part or its content type is invalid.</exception>
        /// <exception cref="InvalidOperationException">The presentation is read-only or closed, or no shape identifier is available.</exception>
        public XSLFChartFrame AddChartFrame(XSLFChart chart)
        {
            if (chart == null)
            {
                throw new ArgumentNullException(nameof(chart));
            }
            EnsureCanEdit();
            var slideShow = GetSlideShow();
            if (!ReferenceEquals(chart.GetPackagePart().Package, GetPackagePart().Package) ||
                slideShow.GetCharts().All(existing => !ReferenceEquals(existing, chart)))
            {
                throw new ArgumentException("The chart must belong to this presentation.", nameof(chart));
            }
            if (chart.GetPackagePart().ContentType != XSLFRelation.CHART.ContentType)
            {
                throw new POIXMLException("The chart part has an invalid content type: " +
                    chart.GetPackagePart().PartName.Name);
            }

            var tree = GetShapeTree();
            if (tree == null)
            {
                throw new POIXMLException("The sheet has no shape tree: " + GetPackagePart().PartName.Name);
            }
            var relationship = FindChartRelationship(chart.GetPackagePart());
            var relationshipId = relationship == null
                ? GetPackagePart().AddRelationship(
                    chart.GetPackagePart().PartName, TargetMode.Internal, XSLFRelation.CHART.Relation).Id
                : relationship.Id;
            var id = AllocateShapeId(tree);
            var frameElement = (XmlElement)tree.AppendChild(CreateFrameElement(tree, id, relationshipId));
            return new XSLFChartFrame(this, frameElement);
        }

        /// <summary>
        /// Returns common slide data of this sheet, or null if the sheet has none.
        /// </summary>
        protected abstract XmlElement GetCommonSlideData();

        private void EnsureCanEdit()
        {
#if NET8_0_OR_GREATER
            ObjectDisposedException.ThrowIf(GetSlideShow().Package == null, this);
#else
            if (GetSlideShow().Package == null)
            {
                throw new ObjectDisposedException(nameof(XSLFSheet));
            }
#endif
            if (GetPackagePart().Package.GetPackageAccess() == PackageAccess.READ)
            {
                throw new InvalidOperationException("Cannot add chart frames in a read-only presentation.");
            }
        }

        internal XmlElement GetShapeTree()
        {
            var commonSlideData = GetCommonSlideData();
            return commonSlideData?["spTree", CT_Presentation.Namespace];
        }

        private void CollectChartFrames(XmlElement tree, List<XSLFChartFrame> frames)
        {
            foreach (XmlNode node in tree.ChildNodes)
            {
                if (!(node is XmlElement element) || element.NamespaceURI != CT_Presentation.Namespace)
                {
                    continue;
                }
                if (element.LocalName == "graphicFrame")
                {
                    var relationshipId = GetChartRelationshipId(element);
                    if (relationshipId != null && IsChartRelationship(relationshipId))
                    {
                        frames.Add(new XSLFChartFrame(this, element));
                    }
                }
                else if (element.LocalName == "grpSp")
                {
                    CollectChartFrames(element, frames);
                }
            }
        }

        private static string GetChartRelationshipId(XmlElement frame)
        {
            var manager = CreateNamespaceManager(frame);
            var reference = (XmlElement)frame.SelectSingleNode(
                "a:graphic/a:graphicData/c:chart[@r:id]", manager);
            var relationshipId = reference?.GetAttribute("id", CT_Presentation.RelationshipsNamespace);
            return string.IsNullOrEmpty(relationshipId) ? null : relationshipId;
        }

        private bool IsChartRelationship(string relationshipId)
        {
            if (!GetPackagePart().HasRelationships)
            {
                return false;
            }
            var relationship = GetPackagePart().GetRelationship(relationshipId);
            if (relationship == null ||
                relationship.RelationshipType != XSLFRelation.CHART.Relation ||
                relationship.TargetMode != TargetMode.Internal)
            {
                return false;
            }
            try
            {
                return GetPackagePart().Package.GetPart(
                    PackagingUriHelper.CreatePartName(relationship.TargetUri))?.ContentType ==
                    XSLFRelation.CHART.ContentType;
            }
            catch (NPOI.OpenXml4Net.Exceptions.InvalidFormatException)
            {
                return false;
            }
        }

        private static XmlNamespaceManager CreateNamespaceManager(XmlElement element)
        {
            var manager = new XmlNamespaceManager(element.OwnerDocument.NameTable);
            manager.AddNamespace("p", CT_Presentation.Namespace);
            manager.AddNamespace("a", XSLFXmlHelper.DrawingNamespace);
            manager.AddNamespace("c", XSLFChart.ChartNamespace);
            manager.AddNamespace("r", CT_Presentation.RelationshipsNamespace);
            return manager;
        }

        private PackageRelationship FindChartRelationship(PackagePart chartPart)
        {
            foreach (PackageRelationship relationship in GetPackagePart().Relationships)
            {
                if (relationship.RelationshipType != XSLFRelation.CHART.Relation ||
                    relationship.TargetMode != TargetMode.Internal)
                {
                    continue;
                }
                try
                {
                    if (ReferenceEquals(GetPackagePart().GetRelatedPart(relationship), chartPart))
                    {
                        return relationship;
                    }
                }
                catch (ArgumentException)
                {
                    // Skip relationships whose target part is missing from the package.
                }
            }
            return null;
        }

        private static uint AllocateShapeId(XmlElement tree)
        {
            uint maxId = 1;
            CollectShapeIds(tree, ref maxId);
            if (maxId >= uint.MaxValue)
            {
                throw new InvalidOperationException("No more shape identifiers are available.");
            }
            return maxId + 1;
        }

        private static void CollectShapeIds(XmlElement element, ref uint maxId)
        {
            foreach (XmlNode node in element.ChildNodes)
            {
                if (!(node is XmlElement child) || child.NamespaceURI != CT_Presentation.Namespace)
                {
                    continue;
                }
                if (child.LocalName == "cNvPr")
                {
                    if (uint.TryParse(child.GetAttribute("id"), NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out var id))
                    {
                        maxId = Math.Max(maxId, id);
                    }
                }
                else
                {
                    CollectShapeIds(child, ref maxId);
                }
            }
        }

        private static XmlElement CreateFrameElement(XmlElement tree, uint id, string relationshipId)
        {
            var namespaces = " xmlns:p=\"" + CT_Presentation.Namespace +
                "\" xmlns:a=\"" + XSLFXmlHelper.DrawingNamespace +
                "\" xmlns:c=\"" + XSLFChart.ChartNamespace +
                "\" xmlns:r=\"" + CT_Presentation.RelationshipsNamespace + "\"";
            var name = "Chart " + id.ToString(CultureInfo.InvariantCulture);
            var frame = XSLFXmlHelper.CreateElement(
                "<p:graphicFrame" + namespaces + ">" +
                "<p:nvGraphicFramePr><p:cNvPr id=\"" +
                id.ToString(CultureInfo.InvariantCulture) + "\" name=\"" + name + "\"/>" +
                "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>" +
                "<p:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></p:xfrm>" +
                "<a:graphic><a:graphicData uri=\"" + XSLFChart.ChartNamespace + "\">" +
                "<c:chart r:id=\"" + relationshipId + "\"/></a:graphicData></a:graphic>" +
                "</p:graphicFrame>");
            return (XmlElement)tree.OwnerDocument.ImportNode(frame, true);
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
    }
}
