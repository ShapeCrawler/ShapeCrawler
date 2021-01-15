using System;
using ShapeCrawler.Models;

namespace ShapeCrawler.Tests.Unit
{
    public class TestFile_018Fixture : IDisposable
    {
        public PresentationSc pre018 { get; }

        public TestFile_018Fixture()
        {
            pre018 = new PresentationSc(Properties.Resources._018);
        }

        public void Dispose()
        {
            pre018.Close();
        }
    }
}