using System.IO;
using ShapeCrawler.Tests.Properties;

namespace ShapeCrawler.Tests.Helpers
{
    public static class TestFiles
    {
        public static class Images
        {
            public static byte[] imageByteArray02 => Resources.test_image_2;
            public static MemoryStream img02_stream => new(Resources.test_image_2);
        }

        public static class Presentations
        {
            public static byte[] pre001 => Resources._001;

            public static Stream pre001_stream => Resources._001.ToResizeableStream();

            public static Stream pre039_stream => Resources._039.ToResizeableStream();

            public static Stream pre011_dt_stream => Resources._011_dt.ToResizeableStream();

            public static byte[] pre025_byteArray => Resources._025;

            public static byte[] pre009 => Resources._009;
        }

        public class Audio
        {
            public static Stream TestMp3 => Resources.test_mp3.ToResizeableStream();
        }
    }
}