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

namespace NPOI.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Util;

    /**
     * A fill color taken from an XFEXT record (MS-XLS 2.4.355). Unlike classic
     * fill colors, which are palette indexes, these colors are stored as raw RGB
     * or as a theme color index with a tint, so several cell styles sharing one
     * legacy palette index may still display different colors.
     */
    public class HSSFExtendedFillColor : HSSFColor
    {
        private readonly FullColorExt color;
        private readonly int legacyIndex;
        private readonly int[][] themeColors;

        /**
         * @param color the extended color stored in the XFEXT record
         * @param legacyIndex the classic palette index kept in the XF record,
         *        used by Excel versions that do not support color extensions
         * @param themeColors RGB triplets of the workbook theme colors
         *        (0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..9=accent1..accent6), or null
         *        if the workbook carries no theme
         */
        public HSSFExtendedFillColor(FullColorExt color, short legacyIndex, int[][] themeColors)
        {
            this.color = color;
            this.legacyIndex = legacyIndex;
            this.themeColors = themeColors;
        }

        /// <summary>
        /// The underlying extended color record
        /// </summary>
        public FullColorExt ColorExt
        {
            get { return color; }
        }

        /// <summary>
        /// The classic palette index stored in the XF record. This is what
        /// Excel versions without color extension support fall back to, and
        /// what older NPOI versions reported.
        /// </summary>
        public override short Indexed
        {
            get { return (short)legacyIndex; }
        }

        /// <summary>
        /// The color type as stored in the XFEXT record:
        /// XCLRAUTO, XCLRINDEXED, XCLRRGB, XCLRTHEMED or XCLRNINCHED
        /// </summary>
        public int ColorType
        {
            get { return color.XclrType; }
        }

        /// <summary>
        /// Theme color index, if <see cref="ColorType"/> is XCLRTHEMED
        /// </summary>
        public int Theme
        {
            get { return (int)color.XclrValue; }
        }

        /// <summary>
        /// Tint of the color, positive lightens, negative darkens
        /// </summary>
        public int TintShade
        {
            get { return color.TintShade; }
        }

        /// <summary>
        /// True if the color type is one this class can resolve to an RGB triplet
        /// </summary>
        public bool IsResolved
        {
            get { return GetTriplet() != null; }
        }

        public override byte[] GetTriplet()
        {
            switch (color.XclrType)
            {
                case XFExtRecord.XCLRAUTO:
                    return Automatic.GetInstance().GetTriplet();
                case XFExtRecord.XCLRINDEXED:
                    int idx = (int)color.XclrValue;
                    if (idx == HSSFColor.Automatic.Index)
                    {
                        return Automatic.GetInstance().GetTriplet();
                    }
                    HSSFColor indexed = HSSFColor.GetIndexHash()[idx];
                    return indexed == null ? null : indexed.GetTriplet();
                case XFExtRecord.XCLRRGB:
                    long rgba = color.XclrValue;
                    return new byte[]
                    {
                        (byte)((rgba >> 16) & 0xFF),
                        (byte)((rgba >> 8) & 0xFF),
                        (byte)(rgba & 0xFF)
                    };
                case XFExtRecord.XCLRTHEMED:
                    int themeIdx = (int)color.XclrValue;
                    if (themeColors == null || themeIdx < 0 || themeIdx >= themeColors.Length)
                    {
                        return null;
                    }
                    return ApplyTint(themeColors[themeIdx], color.TintShade);
                default:
                    return null;
            }
        }

        public override String GetHexString()
        {
            byte[] triplet = GetTriplet();
            if (triplet == null)
            {
                return base.GetHexString();
            }
            return ToHexString(triplet[0]) + ":" + ToHexString(triplet[1]) + ":" + ToHexString(triplet[2]);
        }

        private static String ToHexString(byte value)
        {
            int c = (value & 0xFF) << 8 | (value & 0xFF);
            return c.ToString("X4");
        }

        /**
         * Applies an MS-XLS nTintShade value to a color, per MS-XLS 2.5.164.
         * Values are positive to lighten, negative to darken.
         */
        internal static byte[] ApplyTint(int[] rgb, int tintShade)
        {
            byte[] result = new byte[rgb.Length];
            for (int i = 0; i < rgb.Length; i++)
            {
                result[i] = ApplyTint(rgb[i] & 0xFF, tintShade);
            }
            return result;
        }

        private static byte ApplyTint(int lum, int tintShade)
        {
            if (tintShade > 0)
            {
                double tint = tintShade / 32767.0;
                return (byte)Math.Round(lum + (255 - lum) * tint);
            }
            if (tintShade < 0)
            {
                double tint = tintShade / 32768.0;
                return (byte)Math.Round(lum * (1 + tint));
            }
            return (byte)lum;
        }
    }
}
