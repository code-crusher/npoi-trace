
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


namespace NPOI.HSSF.Record
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using NPOI.Util;

    /**
     * XFEXT Record - Extended XF record (FRT with rt=0x087D).
     * <p>
     * Extends the formatting of an XF record with properties that have no room
     * in the classic XF record, most notably the fill colors which, in newer
     * files, are stored as (theme, tint) or raw RGB values instead of a palette
     * index.
     * <p>
     * Structure per MS-XLS 2.4.355 (XFEXT):
     * <pre>
     * frtHeader   (12 bytes) - rt(2)=0x087D, grbitFrt(2), reserved(8)
     * reserved1   (2 bytes)
     * ixfe        (2 bytes) - index of the XF record being extended
     * reserved2   (2 bytes)
     * cexts       (2 bytes) - number of elements in rgExt
     * rgExt       (variable) - sequence of ExtProp: extType(2), cb(2), extPropData(cb-4)
     * </pre>
     * Only the cell interior foreground/background color extensions
     * (extType 0x0004 / 0x0005, FullColorExt) are interpreted; everything
     * else is kept as raw bytes and round-tripped on save.
     */
    public class XFExtRecord : StandardRecord
    {
        public const short sid = unchecked((short)0x087D);

        /// <summary>extType: cell interior foreground color (FullColorExt)</summary>
        public const int EXTPROP_INTERIOR_FOREGROUND = 0x0004;
        /// <summary>extType: cell interior background color (FullColorExt)</summary>
        public const int EXTPROP_INTERIOR_BACKGROUND = 0x0005;

        /** XColorType: automatic color */
        public const int XCLRAUTO = 0x0;
        /** XColorType: indexed color, xclrValue = IcvXF */
        public const int XCLRINDEXED = 0x1;
        /** XColorType: RGBA color, xclrValue = LongRGBA */
        public const int XCLRRGB = 0x2;
        /** XColorType: theme color, xclrValue = ColorTheme */
        public const int XCLRTHEMED = 0x3;
        /** XColorType: color not set */
        public const int XCLRNINCHED = 0x4;

        private int field_1_ixfe;
        private List<ExtProp> field_2_exts = new List<ExtProp>();

        public XFExtRecord()
        {
        }

        public XFExtRecord(RecordInputStream in1)
        {
            // frtHeader: rt(2), grbitFrt(2), reserved(8)
            in1.ReadShort();
            in1.ReadShort();
            in1.ReadLong();   // reserved, MUST be zero
            in1.ReadShort();  // reserved1

            field_1_ixfe = in1.ReadUShort();
            in1.ReadShort(); // reserved2

            int cexts = in1.ReadUShort();
            for (int i = 0; i < cexts; i++)
            {
                field_2_exts.Add(new ExtProp(in1));
            }
        }

        /**
         * @return the index of the XF record this record extends
         */
        public int Ixfe
        {
            get { return field_1_ixfe; }
            set { field_1_ixfe = value; }
        }

        public IList<ExtProp> ExtProps
        {
            get { return field_2_exts; }
        }

        /**
         * Returns the raw ExtPropData of the given extension type, or null if absent
         */
        public byte[] GetExtPropData(int extType)
        {
            foreach (ExtProp p in field_2_exts)
            {
                if (p.ExtType == extType)
                {
                    return p.Data;
                }
            }
            return null;
        }

        /**
         * Sets (replaces or appends) the given extension property
         */
        public void SetExtPropData(int extType, byte[] data)
        {
            for (int i = 0; i < field_2_exts.Count; i++)
            {
                if (field_2_exts[i].ExtType == extType)
                {
                    field_2_exts[i] = new ExtProp(extType, data);
                    return;
                }
            }
            field_2_exts.Add(new ExtProp(extType, data));
        }

        /**
         * Returns the FullColorExt payload of the given color extension type
         * (interior foreground or background), or null if this record does not
         * carry such an extension.
         */
        public FullColorExt GetColorExt(int colorExtType)
        {
            byte[] data = GetExtPropData(colorExtType);
            if (data != null && data.Length >= 8)
            {
                return new FullColorExt(data);
            }
            return null;
        }

        /**
         * Sets the FullColorExt payload of the given color extension type
         */
        public void SetColorExt(int colorExtType, FullColorExt color)
        {
            SetExtPropData(colorExtType, color.Serialize());
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            // frtHeader
            out1.WriteShort(sid & 0xFFFF); // rt
            out1.WriteShort(0);            // grbitFrt
            out1.WriteInt(0);              // reserved
            out1.WriteInt(0);              // reserved
            // reserved1
            out1.WriteShort(0);
            out1.WriteShort(field_1_ixfe);
            out1.WriteShort(0);            // reserved2
            out1.WriteShort(field_2_exts.Count);
            foreach (ExtProp p in field_2_exts)
            {
                p.Serialize(out1);
            }
        }

        protected override int DataSize
        {
            get
            {
                int size = 12 + 2 + 2 + 2 + 2;
                foreach (ExtProp p in field_2_exts)
                {
                    size += p.EncodedSize;
                }
                return size;
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[XFEXT]\n");
            buffer.Append("    .ixfe    = ")
                .Append(StringUtil.ToHexString(field_1_ixfe)).Append("\n");
            buffer.Append("    .cexts   = ").Append(field_2_exts.Count).Append("\n");
            foreach (ExtProp p in field_2_exts)
            {
                buffer.Append("    .extType = ")
                    .Append(StringUtil.ToHexString(p.ExtType)).Append("\n");
                buffer.Append(p.Data == null ? "    .data    = (null)\n" :
                    HexDump.Dump(p.Data, 0, 0));
            }
            buffer.Append("[/XFEXT]\n");
            return buffer.ToString();
        }

        public override short Sid
        {
            get { return sid; }
        }

        /**
         * ExtProp - a single formatting property extension per MS-XLS 2.5.108.
         * cb counts the size of the whole structure, including its own 4 byte header.
         */
        public class ExtProp
        {
            private int extType;
            private byte[] data;

            public ExtProp(int extType, byte[] data)
            {
                this.extType = extType;
                this.data = data;
            }

            public ExtProp(RecordInputStream in1)
            {
                extType = in1.ReadUShort();
                int cb = in1.ReadUShort(); // size of the whole ExtProp structure
                int dataSize = cb - 4;
                if (dataSize < 0)
                {
                    throw new RecordFormatException("Bad ExtProp size " + cb);
                }
                data = new byte[dataSize];
                if (dataSize > 0)
                {
                    in1.ReadFully(data);
                }
            }

            public int ExtType
            {
                get { return extType; }
            }

            public byte[] Data
            {
                get { return data; }
            }

            public int EncodedSize
            {
                get { return 4 + (data == null ? 0 : data.Length); }
            }

            public void Serialize(ILittleEndianOutput out1)
            {
                out1.WriteShort(extType);
                out1.WriteShort(EncodedSize);
                if (data != null && data.Length > 0)
                {
                    out1.Write(data);
                }
            }
        }
    }

    /**
     * FullColorExt - color payload of an ExtProp per MS-XLS 2.5.155.
     * <pre>
     * xclrType   (2 bytes) - XColorType: 0=auto, 1=indexed, 2=RGB, 3=theme, 4=none
     * nTintShade (2 bytes) - signed tint; positive lightens, negative darkens
     * xclrValue  (4 bytes) - color data; meaning depends on xclrType:
     *                          1: IcvXF (palette index)
     *                          2: LongRGBA (0xAARRGGBB)
     *                          3: ColorTheme (theme color index)
     * </pre>
     */
    public class FullColorExt
    {
        private int field_1_xclrType;
        private int field_2_tintShade;
        private long field_3_xclrValue;

        public FullColorExt()
        {
        }

        public FullColorExt(int xclrType, int tintShade, long xclrValue)
        {
            this.field_1_xclrType = xclrType;
            this.field_2_tintShade = tintShade;
            this.field_3_xclrValue = xclrValue;
        }

        public FullColorExt(byte[] data)
        {
            field_1_xclrType = LittleEndian.GetUShort(data, 0);
            field_2_tintShade = LittleEndian.GetShort(data, 2);
            field_3_xclrValue = LittleEndian.GetUInt(data, 4);
        }

        /// <summary>
        /// XColorType of this color
        /// </summary>
        public int XclrType
        {
            get { return field_1_xclrType; }
            set { field_1_xclrType = value; }
        }

        /// <summary>
        /// Tint of the color, positive lightens, negative darkens (MS-XLS 2.5.164 nTintShade)
        /// </summary>
        public int TintShade
        {
            get { return field_2_tintShade; }
            set { field_2_tintShade = value; }
        }

        /// <summary>
        /// Color value; indexed palette index, 0xAARRGGBB RGBA or theme index depending on XclrType
        /// </summary>
        public long XclrValue
        {
            get { return field_3_xclrValue; }
            set { field_3_xclrValue = value; }
        }

        public byte[] Serialize()
        {
            byte[] data = new byte[8];
            LittleEndian.PutUShort(data, 0, field_1_xclrType);
            LittleEndian.PutShort(data, 2, (short)field_2_tintShade);
            LittleEndian.PutUInt(data, 4, field_3_xclrValue);
            return data;
        }

        public override String ToString()
        {
            return "[FullColorExt] type=" + XclrType + " tint=" + TintShade
                + " value=" + StringUtil.ToHexString(XclrValue);
        }
    }
}
