using NPOI.Util;
using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace TestCases.Util
{
    internal class TestPngUtils
    {
        [Test]
        public void TestMatchesPngHeader()
        {
            byte[] pngData = new byte[] { (byte)0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
            ClassicAssert.IsTrue(PngUtils.MatchesPngHeader(pngData, 0));

            byte[] nonPngData = new byte[] { 0x00, 0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77 };
            ClassicAssert.IsFalse(PngUtils.MatchesPngHeader(nonPngData, 0));

            byte[] dataWithOffset = new byte[] { 0x00, (byte)0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
            ClassicAssert.IsTrue(PngUtils.MatchesPngHeader(dataWithOffset, 1));
        }
    }
}
