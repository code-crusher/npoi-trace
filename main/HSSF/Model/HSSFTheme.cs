/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSSF.Model
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;
    using ICSharpCode.SharpZipLib.Zip;
    using NPOI.HSSF.Record;
    using NPOI.Util;

    /**
     * The theme colors of a .xls workbook. Excel 2007 and later embed the theme
     * of the workbook as an OOXML package inside a THEMEDATA (0x0896) record of
     * the workbook stream. Fill and font colors may reference these theme colors
     * through XFEXT records.
     */
    public class HSSFTheme
    {
        /**
         * sid of the THEMEDATA record, an FRT record carrying an embedded OOXML package
         */
        public const short SID_THEMEDATA = unchecked((short)0x0896);

        /**
         * Number of theme colors addressable through XFEXT records:
         * lt1, dk1, lt2, dk2, accent1-6
         */
        private const int THEME_COLOR_COUNT = 10;

        private static readonly string[] THEME_ELEMENT_NAMES =
        {
            "lt1", "dk1", "lt2", "dk2",
            "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"
        };

        private static readonly string[] THEME_PART_NAMES =
        {
            "theme/theme/theme1.xml",
            "xl/theme/theme1.xml",
            "theme/theme1.xml"
        };

        private readonly int[][] themeColors;

        public HSSFTheme(int[][] themeColors)
        {
            this.themeColors = themeColors;
        }

        /**
         * @return the RGB triplets of the theme colors indexed like XFEXT theme
         *         references (0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..9=accent1..6), or
         *         null where a color could not be determined
         */
        public int[][] GetThemeColors()
        {
            return themeColors;
        }

        /**
         * Extracts the theme colors from the workbook record list, or returns
         * null if the workbook carries no readable THEMEDATA record.
         */
        public static HSSFTheme FromWorkbook(IList<Record> records)
        {
            foreach (Record rec in records)
            {
                if (rec.Sid == SID_THEMEDATA)
                {
                    byte[] data = rec.Serialize();
                    if (data.Length > 16)
                    {
                        using (MemoryStream stream = new MemoryStream(data, 16, data.Length - 16, false))
                        {
                            return new HSSFTheme(ParseThemeXml(ReadThemePart(stream)));
                        }
                    }
                }
            }
            return null;
        }

        private static byte[] ReadThemePart(Stream package)
        {
            try
            {
                using (ZipInputStream zip = new ZipInputStream(package))
                {
                    ZipEntry entry;
                    while ((entry = zip.GetNextEntry()) != null)
                    {
                        foreach (string name in THEME_PART_NAMES)
                        {
                            if (string.Equals(entry.Name, name, StringComparison.OrdinalIgnoreCase))
                            {
                                return IOUtils.ToByteArray(zip);
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                // a corrupt or unexpected package must not break opening the workbook
            }
            return null;
        }

        private static int[][] ParseThemeXml(byte[] xml)
        {
            if (xml == null)
            {
                return null;
            }
            try
            {
                XmlDocument doc = new XmlDocument { XmlResolver = null };
                using (var stream = new MemoryStream(xml, false))
                {
                    doc.Load(stream);
                }

                XmlElement scheme = (XmlElement)doc.DocumentElement.SelectSingleNode(
                    "//*[local-name()='clrScheme']");
                if (scheme == null)
                {
                    return null;
                }
                int[][] colors = new int[THEME_COLOR_COUNT][];
                for (int i = 0; i < THEME_COLOR_COUNT; i++)
                {
                    colors[i] = ParseColor(scheme, THEME_ELEMENT_NAMES[i]);
                }
                return colors;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int[] ParseColor(XmlElement scheme, string elementName)
        {
            foreach (XmlNode node in scheme.ChildNodes)
            {
                if (node.NodeType != XmlNodeType.Element ||
                    !string.Equals(node.LocalName, elementName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                foreach (XmlNode child in node.ChildNodes)
                {
                    if (child.NodeType != XmlNodeType.Element)
                    {
                        continue;
                    }
                    if (string.Equals(child.LocalName, "srgbClr", StringComparison.OrdinalIgnoreCase))
                    {
                        return ParseHex(((XmlElement)child).GetAttribute("val"));
                    }
                    if (string.Equals(child.LocalName, "sysClr", StringComparison.OrdinalIgnoreCase))
                    {
                        string lastClr = ((XmlElement)child).GetAttribute("lastClr");
                        if (lastClr != null && lastClr.Length > 0)
                        {
                            return ParseHex(lastClr);
                        }
                    }
                }
                return null;
            }
            return null;
        }

        private static int[] ParseHex(string hex)
        {
            if (hex == null || hex.Length < 6)
            {
                return null;
            }
            return new int[]
            {
                Convert.ToInt32(hex.Substring(0, 2), 16),
                Convert.ToInt32(hex.Substring(2, 2), 16),
                Convert.ToInt32(hex.Substring(4, 2), 16)
            };
        }
    }
}
