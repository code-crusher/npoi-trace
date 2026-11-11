using System.IO;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Chart;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A read-only classic DrawingML chart part. The original part bytes are preserved on save;
    /// XML and model snapshots are detached copies that never write back to the package.
    /// </summary>
    public class XSLFChart : POIXMLDocumentPart
    {
        /// <summary>
        /// The classic DrawingML chart namespace used by graphic frame references.
        /// </summary>
        public const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        internal XSLFChart(POIXMLDocumentPart parent, PackagePart part)
            : base(parent, part)
        {
            Committed = true;
        }

        /// <summary>
        /// Returns the original bytes of the chart part.
        /// </summary>
        public byte[] GetData()
        {
            using var input = GetPackagePart().GetInputStream();
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }

        /// <summary>
        /// Returns a detached XML document snapshot of the chart part.
        /// Changes to the snapshot do not affect the presentation.
        /// </summary>
        /// <exception cref="POIXMLException">The chart part is not well-formed XML.</exception>
        public XmlDocument GetXmlSnapshot()
        {
            var document = new XmlDocument { XmlResolver = null };
            using (var input = GetPackagePart().GetInputStream())
            {
                using var reader = XmlReader.Create(input, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null
                });
                try
                {
                    document.Load(reader);
                }
                catch (XmlException e)
                {
                    throw new POIXMLException("Invalid DrawingML chart: " + GetPackagePart().PartName.Name, e);
                }
            }
            return document;
        }

        /// <summary>
        /// Returns a detached chart-space model snapshot parsed with the NPOI chart schema model.
        /// Changes to the snapshot do not affect the presentation.
        /// </summary>
        /// <exception cref="POIXMLException">The chart part is not valid classic DrawingML.</exception>
        public CT_ChartSpace GetChartSnapshot()
        {
            var document = GetXmlSnapshot();
            var root = document.DocumentElement;
            if (root == null || root.LocalName != "chartSpace" || root.NamespaceURI != ChartNamespace)
            {
                throw new POIXMLException("Invalid DrawingML chart: " + GetPackagePart().PartName.Name);
            }
            return ChartSpaceDocument.Parse(document, NamespaceManager).GetChartSpace();
        }
    }
}
