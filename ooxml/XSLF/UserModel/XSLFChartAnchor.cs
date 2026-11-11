namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// An explicit point-based anchor in English Metric Units (EMU).
    /// </summary>
    public struct XSLFChartAnchor
    {
        /// <summary>
        /// Creates an anchor with the given position and size in EMUs.
        /// </summary>
        public XSLFChartAnchor(long x, long y, long width, long height)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        /// <summary>
        /// Gets the horizontal position in EMUs.
        /// </summary>
        public long X { get; }

        /// <summary>
        /// Gets the vertical position in EMUs.
        /// </summary>
        public long Y { get; }

        /// <summary>
        /// Gets the width in EMUs.
        /// </summary>
        public long Width { get; }

        /// <summary>
        /// Gets the height in EMUs.
        /// </summary>
        public long Height { get; }
    }
}
