using System;
using ShapeCrawler.Models;

namespace ShapeCrawler.Tests.Unit
{
    public class TestFile_018Fixture : IDisposable
    {
        public PresentationEx pre018 { get; }

        public TestFile_018Fixture()
        {
            pre018 = new PresentationEx(Properties.Resources._018);
        }

        public void Dispose()
        {
            pre018.Close();
        }
    }
}