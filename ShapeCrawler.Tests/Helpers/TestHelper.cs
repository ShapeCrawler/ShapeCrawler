using System.IO;

namespace ShapeCrawler.Tests.Helpers
{
    public static class TestHelper
    {
        static TestHelper()
        {
            HorizontalResolution = 96;
            VerticalResolution = 96;
        }

        public static MemoryStream ToResizeableStream(this byte[] byteArray)
        {
            var stream = new MemoryStream();
            stream.Write(byteArray, 0, byteArray.Length);

            return stream;
        }

        public static readonly float HorizontalResolution;
        
        public static readonly float VerticalResolution;
    }
}