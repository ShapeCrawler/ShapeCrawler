using System;
using System.IO;
using ShapeCrawler.Models;
using SlideDotNet.Models;

namespace ShapeCrawler.UnitTests
{
    public class TestFile_002Fixture : IDisposable
    {
        public PresentationEx pre002 { get; }

        public TestFile_002Fixture()
        {
            var ms = new MemoryStream(Properties.Resources._002);
            pre002 = new PresentationEx(ms);
        }

        public void Dispose()
        {
            pre002.Close();
        }
    }
}