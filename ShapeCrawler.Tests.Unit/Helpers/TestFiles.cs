using System.IO;
using ShapeCrawler.Tests.Unit.Properties;

namespace ShapeCrawler.Tests.Unit
{
    public static class TestFiles
    {
        public static class Images
        {
            public static byte[] img02 => Resources.test_image_2;
            public static MemoryStream img02_stream => new(Resources.test_image_2);
        }

        public static class Presentations
        {
            public static byte[] pre009 => Resources._009;
        }
    }
}