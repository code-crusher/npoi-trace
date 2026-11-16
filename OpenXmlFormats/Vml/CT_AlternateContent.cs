using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Vml
{
    public class CT_AlternateContent
    {
        private const string MarkupCompatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        public string InnerXml { get; set; }
        private readonly Dictionary<string, string> namespaces = new Dictionary<string, string>();
        public CT_AlternateContent()
        {
        }
        public static CT_AlternateContent Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;

            var ac = new CT_AlternateContent();
            // preserve namespace declarations (e.g. xmlns:a14) referenced by attribute values
            // such as mc:Requires="a14": they are lost when only the inner xml is kept,
            // which makes the round-tripped AlternateContent invalid
            foreach (XmlAttribute attribute in node.Attributes)
            {
                if (attribute.Prefix == "xmlns" || (attribute.Prefix == string.Empty && attribute.LocalName == "xmlns"))
                {
                    ac.namespaces[attribute.Name] = attribute.Value;
                }
            }
            if (string.IsNullOrEmpty(node.InnerXml))
            {
                return ac;
            }
            ac.InnerXml = node.InnerXml;
            return ac;
        }
        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.WriteStart("mc", nodeName);
            sw.WriteAttribute("xmlns:mc", MarkupCompatibilityNamespace);
            foreach (KeyValuePair<string, string> declaration in this.namespaces)
            {
                if (declaration.Key != "xmlns:mc")
                {
                    XmlHelper.WriteAttribute(sw, declaration.Key, declaration.Value);
                }
            }
            if (this.InnerXml == null)
            {
                sw.Write("/>");
            }
            else
            {
                sw.Write('>');
                sw.Write(this.InnerXml);
                sw.WriteEndElement("mc", nodeName);
            }
        }
    }
}
