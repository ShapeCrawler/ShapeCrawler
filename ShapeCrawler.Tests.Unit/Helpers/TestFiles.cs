using System.IO;
using ShapeCrawler.Tests.Unit.Properties;

namespace ShapeCrawler.Tests.Unit.Helpers
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
            public static byte[] pre001 => Resources._001;

            public static Stream pre001_stream
            {
                get
                {
                    var stream = new MemoryStream();
                    stream.Write(Resources._001, 0, Resources._001.Length);
                    return stream;
                }
            }

            public static Stream pre025_pptxStream
            {
                get
                {
                    var stream = new MemoryStream();
                    stream.Write(Resources._025, 0, Resources._025.Length);
                    return stream;
                }
            }

            public static byte[] pre009 => Resources._009;
        }
    }
}