using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ShapeCrawler.Tests.Helpers
{
    public static class TestHelper
    {
        static TestHelper()
        {
            var bm = new Bitmap(100, 100);
            if (bm.HorizontalResolution == 0)
            {
                // Set default resolution
                bm.SetResolution(96, 96);
            }

            HorizontalResolution = bm.HorizontalResolution;
            VerticalResolution = bm.VerticalResolution;
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