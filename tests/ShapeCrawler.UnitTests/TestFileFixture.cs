using System;
using System.IO;
using ShapeCrawler.Models;

namespace ShapeCrawler.UnitTests
{
    public class TestFileFixture : IDisposable
    {
        public PresentationEx pre002 { get; }

        public TestFileFixture()
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